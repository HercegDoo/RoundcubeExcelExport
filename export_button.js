/**
 * Export Excel Plugin for Roundcube Webmail
 * @type {{init: ExportExcelPlugin.init}}
 */
const ExportExcelPlugin = function () {

    /**
     * @type {string[]} formats - list of available formats to export, defined in the button's data-formats attribute
     */
    let formats;

    /**
     * Registers the menu for the export button and sets up the necessary event listeners and commands.
     */
    const registerMenu = function () {

        // find and modify default download link/button
        $.each(rcmail.buttons['export_excel'] || [], function () {
            var link = $('#' + this.id),
                span = $('span', link);

            if (!span.length) {
                span = $('<span>');
                link.html('').append(span);
            }

            formats = link.data('formats').split(',') || [];

            link.attr('aria-haspopup', 'true');

        });

        // commands status
        rcmail.message_list && rcmail.message_list.addEventListener('select', function (list) {
            var selected = list.get_selection().length;

            rcmail.enable_command('export_excel', selected > 0);

            formats.forEach(function (format) {
                rcmail.enable_command('export_excel_' + format, selected > 0);
            });
        });

        // register command export_excel to open the export-excel-menu menu
        rcmail.register_command('export_excel', function (opt, element, event) {
            rcmail.command('menu-open', 'export-excel-menu', element, event);
        }, true);

    }

    /**
     * Registers the commands for each available format to export.
     */
    const registerCommands = function () {
        formats.forEach(function (format) {
            rcmail.register_command('export_excel_' + format, function () {
                exportHandler(format);
            }, true);
        });
    }

    /**
     * Display loading indicator and set interval to check if the file is cookie removed
     * @param requestId
     */
    const displayLoadingIndicator = function (requestId) {
        const expires = new Date();
        expires.setDate(expires.getDate() + 1);

        rcmail.set_cookie(requestId, 'exporting', expires);

        const lock = rcmail.set_busy(true, 'RoundcubeExcelExport.exporting');

        // set interval to check if the file is cookie removed
        const interval = setInterval(() => {
            if (rcmail.get_cookie(requestId) !== 'exporting') {
                rcmail.set_busy(false, null, lock);
                clearInterval(interval);
                clearTimeout(timeout);
            }
        }, 250);

        // get from $config max_execution_time
        const maxExecutionTime = (rcmail.env.plugin_export_excel_settings.max_execution_time || 120) * 1000;

        // set timeout to 120 seconds and remove cookie then show error
        const timeout = setTimeout(() => {
            rcmail.set_busy(false, null, lock);
            rcmail.display_message('Server Internal Error', 'error');
            clearInterval(interval);
        }, maxExecutionTime);

    }


    /**
     * Export handler, sends a POST request to the server to download the selected messages in the specified format.
     * @param format
     */
    const exportHandler = function (format) {
        // multi-message download, use hidden form to POST selection
        if (rcmail.message_list && rcmail.message_list.get_selection().length > 0) {


            const inputs = [],
                post = rcmail.selection_post_data(),
                id = 'export-excel-' + new Date().getTime(),
                iframe = $('<iframe>').attr({name: id, style: 'display:none'}),
                form = $('<form>').attr({
                    target: id,
                    style: 'display: none',
                    method: 'post',
                    action: rcmail.url('mail/plugin.export.excel')
                });

            post._format = format;
            post._token = rcmail.env.request_token;
            post._request_id = id;

            $.each(post, function (k, v) {
                if (typeof v == 'object' && v.length > 1) {
                    for (let j = 0; j < v.length; j++)
                        inputs.push($('<input>').attr({type: 'hidden', name: k + '[]', value: v[j]}));
                } else {
                    inputs.push($('<input>').attr({type: 'hidden', name: k, value: v}));
                }
            });

            iframe.appendTo(document.body);
            displayLoadingIndicator(id)
            form.append(inputs).appendTo(document.body).submit();
        }
    }

    return {
        init: function () {
            registerMenu();
            registerCommands();
        }
    };
}();

// initialize the plugin
window.rcmail && rcmail.addEventListener('init', function () {
    ExportExcelPlugin.init();
});



