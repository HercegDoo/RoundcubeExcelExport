# Roundcube Excel Export

## Description
**Roundcube Excel Export** is a plugin for the Roundcube webmail client that allows users to export the contents of an email folder to an Excel file (.xlsx). This plugin is designed to help users easily back up or analyze their emails in a structured format.

## Features
- Export email subjects, senders, body and dates to an Excel file.
- Simple integration with the Roundcube interface.
- Lightweight and easy to use.

## Installation

1. **Download or Clone the Repository:**

   ```bash
   git clone https://github.com/your-username/RoundcubeExcelExport.git

2. **Copy the Plugin:**
   Copy the `RoundcubeExcelExport` folder to your Roundcube installation's `plugins` directory.

   ```bash
   cp -r RoundcubeExcelExport /path/to/roundcube/plugins/
   ```

3. **Enable the Plugin:**
   Edit your Roundcube `config/config.inc.php` file to enable the plugin by adding it to the `$config['plugins']` array.

   ```php
   $config['plugins'] = array('RoundcubeExcelExport', ...);
   ```

4. **Install Dependencies:**
   You can install the required dependencies using Composer. Navigate to the plugin directory and run the following command:

   ```bash
   composer install
   ```

## Usage

1. Log in to your Roundcube webmail client.
2. Navigate to the email folder you wish to export.
3. Click the "Export to Excel" button in the toolbar.
4. The contents of the folder will be downloaded as an Excel file.

## Requirements
- Roundcube 1.3 or higher
- PHP 7.0 or higher
- PHPExcel or PhpSpreadsheet library

## License
This plugin is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.

## Contributing
Contributions are welcome! Please fork this repository and submit a pull request for any features, bug fixes, or enhancements.

## Contact
For issues or questions, please open an issue on GitHub or contact the maintainer at [rti@dooherceg.ba](rti@dooherceg.ba).

