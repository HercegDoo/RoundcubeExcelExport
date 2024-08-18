<?php


require_once __DIR__ . DIRECTORY_SEPARATOR . implode(DIRECTORY_SEPARATOR, ['vendor', 'autoload.php']);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Html as HtmlWriter;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

final class rc_excel_export extends rcube_plugin
{
    /**
     * @var string $task
     */
    public $task = 'mail';

    /**
     * Array mapping export format extensions to their corresponding writer classes.
     *
     * @var array{string: class-string<IWriter>}
     */
    private array $exportFormats = [
        'xlsx' => Xlsx::class,
        'xls' => Xls::class,
        'csv' => Csv::class,
        'html' => HtmlWriter::class,
    ];
    private string $charset = 'ASCII';

    /**
     * Initializes the plugin by preparing the configuration, adding localization texts,
     * registering hooks, and setting up the export action.
     *
     * @return void
     */
    public function init(): void
    {
        $this->prepare_config();
        $this->add_texts('localization');
        $this->add_hook('startup', [$this, 'startup']);
        $this->register_action('plugin.export.excel', [$this, 'exportToExcel']);


        $rcmail = rcmail::get_instance();

        if (!$rcmail->action) {
            $this->exportMenu();
        }
    }

    /**
     * Prepares the configuration settings by loading the configuration and setting the charset.
     *
     * @return void
     */
    private function prepare_config(): void
    {
        // Get the instance of the rcmail class
        $rcmail = rcmail::get_instance();

        // Load the configuration settings
        $this->load_config();

        // Set the charset for the Excel export, defaulting to RCUBE_CHARSET if not specified
        $this->charset = $rcmail->config->get('excel_charset', RCUBE_CHARSET);
    }

    /**
     * Adds the export button to the message menu and the export menu to the footer.
     *
     * @return void
     */
    public function exportMenu()
    {
        $this->include_script('export_button.js');


        $this->api->add_content(
            html::tag(
                'li',
                ['role' => 'menuitem'],
                $this->api->output->button([
                    'command' => 'export_excel',
                    'label' => 'rc_excel_export.export_excel',
                    'type' => 'link',
                    'classact' => 'icon export active',
                    'class' => 'icon export disabled',
                    'innerclass' => 'icon export',
                    'data-formats' => implode(',', $this->formats())
                ])
            ),
            'messagemenu'
        );


        $rcmail = rcmail::get_instance();
        $menu = [];
        $ul_attr = [
            'role' => 'menu',
            'aria-labelledby' => 'aria-label-exportexcelmenu',
            'class' => 'toolbarmenu menu',
        ];

        foreach ($this->formats() as $type) {
            $menu[] = html::tag('li', null, $rcmail->output->button([
                'command' => "export_excel_{$type}",
                'label' => "rc_excel_export.export_excel_{$type}",
                'class' => "download {$type} disabled",
                'classact' => "download {$type} active",
                'type' => 'link',
            ]));
        }

        $rcmail->output->add_footer(
            html::div(['id' => 'export-excel-menu', 'class' => 'popupmenu', 'aria-hidden' => 'true'],
                html::tag('h2', ['class' => 'voice', 'id' => 'aria-label-exportexcelmenu'], "Message Export Options Menu")
                . html::tag('ul', $ul_attr, implode('', $menu))
            )
        );
    }

    /**
     * Returns the list of supported export formats.
     *
     * @return string[]
     */
    private function formats(): array
    {
        return array_keys($this->exportFormats);
    }

    /**
     * Adds the export button to the message menu and the export menu to the footer.
     *
     * @return void
     */
    public function startup(): void
    {
        $rcmail = rcmail::get_instance();
        $rcmail->output->set_env('plugin_export_excel_settings', [
            'max_execution_time' => $rcmail->config->get('max_execution_time', 900),
        ]);
        $rcmail->output->add_label('rc_excel_export.exporting');
    }

    /**
     * Exports the selected messages to an Excel file in the specified format.
     *
     * @return void
     */
    public function exportToExcel()
    {
        $this->removeLoadingSpinner();
        $rcmail = rcmail::get_instance();
        $messageset = rcmail_action::get_uids(null, null, $multi, rcube_utils::INPUT_POST);
        $format = rcube_utils::get_input_string('_format', rcube_utils::INPUT_POST) ?? 'xlsx';

        // Validate the format
        if (!array_key_exists($format, $this->exportFormats)) {
            $msg = $this->gettext([
                'name' => 'wrong_format',
                'vars' => ['$format' => $format]
            ]);

            $rcmail->output->show_message($msg, 'error');
            $rcmail->output->send('iframe');
        }

        if (count($messageset) > 0) {

            // set memory limit to 2048M and exicution time to 900
            if ($rcmail->config->get('memory_limit')) {
                @ini_set('memory_limit', $rcmail->config->get('memory_limit'));
            }
            if ($rcmail->config->get('max_execution_time')) {
                @ini_set('max_execution_time', $rcmail->config->get('max_execution_time'));
            }

            $this->generateExcelFile($messageset, $format);
        }
    }

    /**
     * Removes the loading spinner from the page by deleting the cookie with the request ID.
     *
     * @return void
     */
    private function removeLoadingSpinner()
    {
        $requestId = rcube_utils::get_input_string('_request_id', rcube_utils::INPUT_POST) ?? null;

        // remove cookie with name of $requestId
        if ($requestId) {
            setcookie($requestId, '', time() - 3600, '/');
        }
    }

    /**
     * Generates an Excel file containing the selected messages.
     *
     * @param array<string, numeric-string[]> $messageSet
     * @param string $format
     * @return void
     * @throws Exception
     */
    private function generateExcelFile(array $messageSet, string $format): void
    {
        $rcmail = rcmail::get_instance();
        $imap = $rcmail->get_storage();

        $spreadsheet = new Spreadsheet();
        $totalSelectedMessages = 0;

        foreach ($messageSet as $mbox => $uids) {
            $imap->set_folder($mbox);


            if ($uids === '*') {
                $index = $imap->index($mbox, null, null, true);
                $uids = $index->get();
            }

            $totalSelectedMessages += count($uids);


            if ($totalSelectedMessages > $rcmail->config->get('limit_number_of_messages', 1000)) {
                $msg = $this->gettext([
                    'name' => 'too_many_messages',
                    'vars' => ['$limit' => $rcmail->config->get('limit_number_of_messages', 1000)]
                ]);
                $rcmail->output->show_message($msg, 'error');
                $rcmail->output->send('iframe');
                exit;
            }


            // create a new sheet for $mbox and set it as active
            $activeWorksheet = $spreadsheet->createSheet();
            // limit 31 chars
            $activeWorksheet->setTitle(substr($mbox, 0, 31));


            // set from,to,subject,date,body to excel header
            $activeWorksheet->setCellValue('A1', $this->gettext('column_from'));
            $activeWorksheet->setCellValue('B1', $this->gettext('column_to'));
            $activeWorksheet->setCellValue('C1', $this->gettext('column_subject'));
            $activeWorksheet->setCellValue('D1', $this->gettext('column_date'));
            $activeWorksheet->setCellValue('E1', $this->gettext('column_body'));

            $row = 2;
            foreach ($uids as $uid) {
                $headers = $imap->get_message_headers($uid);
                $body = $imap->get_body($uid);

                // Detect and convert the charset properly
                $charset = $headers->charset ?: RCUBE_CHARSET;
                $subject = rcube_charset::convert($headers->subject, $charset, $this->charset);

                $body = quoted_printable_decode($body);
                $body = rcube_charset::convert($body, $charset, $this->charset);
                $body = $rcmail->html2text($body);

                $activeWorksheet->setCellValue('A' . $row, $headers->from);
                $activeWorksheet->setCellValue('B' . $row, $headers->to);
                $activeWorksheet->setCellValue('C' . $row, $subject);
                $activeWorksheet->setCellValue('D' . $row, $headers->date);
                $activeWorksheet->setCellValue('E' . $row, "'{$body}'");

                $row++;
            }
        }

        // remove the default sheet created by PhpSpreadsheet
        $spreadsheet->removeSheetByIndex(0);

        // set the headers
        $rcmail->output->download_headers(
            "ExcelExport_" . date('Y-m-d_H-i-s') . ".$format", [
            'type_charset' => $this->charset,
            'time_limit' => $rcmail->config->get('max_execution_time', (int)ini_get('max_execution_time')) + 30,
        ]);

        /** @var IWriter $writer */
        $writer = new $this->exportFormats[$format]($spreadsheet);

        if ($writer instanceof Csv) {
            // set the CSV writer settings
            $writer->setUseBOM(true); // Add BOM for UTF-8 encoding
            $writer->setDelimiter(",");
            $writer->setEnclosure('"');
            $writer->setLineEnding("\r\n");
            $writer->setSheetIndex(0);
        }

        $writer->save('php://output');
        exit;
    }


}