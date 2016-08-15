<?php

/**
 * Define the internationalization functionality
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @link       http://sirdar.fi
 * @since      1.0.0
 *
 * @package    Gravity_Forms_Export_Entries_Excel
 * @subpackage Gravity_Forms_Export_Entries_Excel/includes
 */

/**
 * Define the internationalization functionality.
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @since      1.0.0
 * @package    Gravity_Forms_Export_Entries_Excel
 * @subpackage Gravity_Forms_Export_Entries_Excel/includes
 * @author     Jukka Rautanen <support@sirdar.fi>
 */
class Gravity_Forms_Export_Entries_Excel_i18n {


	/**
	 * Load the plugin text domain for translation.
	 *
	 * @since    1.0.0
	 */
	public function load_plugin_textdomain() {

		load_plugin_textdomain(
			'gravity-forms-export-entries-excel',
			false,
			dirname( dirname( plugin_basename( __FILE__ ) ) ) . '/languages/'
		);

	}



}
