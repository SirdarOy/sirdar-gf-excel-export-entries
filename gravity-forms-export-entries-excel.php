<?php

/**
 * The plugin bootstrap file
 *
 * This file is read by WordPress to generate the plugin information in the plugin
 * admin area. This file also includes all of the dependencies used by the plugin,
 * registers the activation and deactivation functions, and defines a function
 * that starts the plugin.
 *
 * @link              http://sirdar.fi
 * @since             1.0.0
 * @package           Gravity_Forms_Export_Entries_Excel
 *
 * @wordpress-plugin
 * Plugin Name:       Gravity Forms Export Entries to Excel
 * Plugin URI:        http://sirdar.fi
 * Description:       This plugin adds Excel export for entries in Gravity Forms
 * Version:           1.0.0
 * Author:            Jukka Rautanen
 * Author URI:        http://sirdar.fi
 * License:           GPL-2.0+
 * License URI:       http://www.gnu.org/licenses/gpl-2.0.txt
 * Text Domain:       gravity-forms-export-entries-excel
 * Domain Path:       /languages
 */

// If this file is called directly, abort.
if ( ! defined( 'WPINC' ) ) {
	die;
}

/**
 * The code that runs during plugin activation.
 * This action is documented in includes/class-gravity-forms-export-entries-excel-activator.php
 */
function activate_gravity_forms_export_entries_excel() {
	require_once plugin_dir_path( __FILE__ ) . 'includes/class-gravity-forms-export-entries-excel-activator.php';
	Gravity_Forms_Export_Entries_Excel_Activator::activate();
}

/**
 * The code that runs during plugin deactivation.
 * This action is documented in includes/class-gravity-forms-export-entries-excel-deactivator.php
 */
function deactivate_gravity_forms_export_entries_excel() {
	require_once plugin_dir_path( __FILE__ ) . 'includes/class-gravity-forms-export-entries-excel-deactivator.php';
	Gravity_Forms_Export_Entries_Excel_Deactivator::deactivate();
}

register_activation_hook( __FILE__, 'activate_gravity_forms_export_entries_excel' );
register_deactivation_hook( __FILE__, 'deactivate_gravity_forms_export_entries_excel' );

/**
 * The core plugin class that is used to define internationalization,
 * admin-specific hooks, and public-facing site hooks.
 */
require plugin_dir_path( __FILE__ ) . 'includes/class-gravity-forms-export-entries-excel.php';

/**
 * Begins execution of the plugin.
 *
 * Since everything within the plugin is registered via hooks,
 * then kicking off the plugin from this point in the file does
 * not affect the page life cycle.
 *
 * @since    1.0.0
 */
function run_gravity_forms_export_entries_excel() {

	$plugin = new Gravity_Forms_Export_Entries_Excel();
	$plugin->run();

}
run_gravity_forms_export_entries_excel();