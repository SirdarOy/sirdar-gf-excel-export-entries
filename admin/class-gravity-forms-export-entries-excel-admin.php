<?php

/**
 * The admin-specific functionality of the plugin.
 *
 * @link       http://sirdar.fi
 * @since      1.0.0
 *
 * @package    Gravity_Forms_Export_Entries_Excel
 * @subpackage Gravity_Forms_Export_Entries_Excel/admin
 */

/**
 * The admin-specific functionality of the plugin.
 *
 * Defines the plugin name, version, and two examples hooks for how to
 * enqueue the admin-specific stylesheet and JavaScript.
 *
 * @package    Gravity_Forms_Export_Entries_Excel
 * @subpackage Gravity_Forms_Export_Entries_Excel/admin
 * @author     Jukka Rautanen <support@sirdar.fi>
 */
class Gravity_Forms_Export_Entries_Excel_Admin {

	/**
	 * The ID of this plugin.
	 *
	 * @since    1.0.0
	 * @access   private
	 * @var      string    $plugin_name    The ID of this plugin.
	 */
	private $plugin_name;

	/**
	 * The version of this plugin.
	 *
	 * @since    1.0.0
	 * @access   private
	 * @var      string    $version    The current version of this plugin.
	 */
	private $version;

	/**
	 * Initialize the class and set its properties.
	 *
	 * @since    1.0.0
	 * @param      string    $plugin_name       The name of this plugin.
	 * @param      string    $version    The version of this plugin.
	 */
	public function __construct( $plugin_name, $version ) {

		$this->plugin_name = $plugin_name;
		$this->version = $version;
		$this->add_export_page();

	}

	public function add_export_page() {
		// add custom menu item
		add_filter( 'gform_export_menu', 'my_custom_export_menu_item' );
		function my_custom_export_menu_item( $menu_items ) {
		    
		    $menu_items[] = array(
		        'name' => 'export_entries_to_excel',
		        'label' => __( 'Export Entries to Excel' )
		        );
		    
		    return $menu_items;
		}

		// display content for custom menu item when selected
		add_action( 'gform_export_page_export_entries_to_excel', 'export_entries_to_excel' );
		function export_entries_to_excel() {
				if ( ! GFCommon::current_user_can_any( 'gravityforms_export_entries' ) ) {
					wp_die( 'You do not have permission to access this page' );
				}


				GFExport::page_header( __( 'Export Entries', 'gravityforms' ) );

				?>

				<script type="text/javascript">

					var gfSpinner;

					<?php GFCommon::gf_global(); ?>
					<?php GFCommon::gf_vars(); ?>

					function SelectExportForm(formId) {

						if (!formId)
							return;

						gfSpinner = new gfAjaxSpinner(jQuery('select#export_form'), gf_vars.baseUrl + '/images/spinner.gif', 'position: relative; top: 2px; left: 5px;');

						var mysack = new sack("<?php echo admin_url( 'admin-ajax.php' )?>");
						mysack.execute = 1;
						mysack.method = 'POST';
						mysack.setVar("action", "rg_select_export_form");
						mysack.setVar("rg_select_export_form", "<?php echo wp_create_nonce( 'rg_select_export_form' ); ?>");
						mysack.setVar("form_id", formId);
						mysack.onError = function () {
							alert(<?php echo json_encode( __( 'Ajax error while selecting a form', 'gravityforms' ) ); ?>)
						};
						mysack.runAJAX();

						return true;
					}

					function EndSelectExportForm(aryFields, filterSettings) {

						gfSpinner.destroy();

						if (aryFields.length == 0) {
							jQuery("#export_field_container, #export_date_container, #export_submit_container").hide()
							return;
						}

						var fieldList = "<li><input id='select_all' type='checkbox' onclick=\"jQuery('.gform_export_field').attr('checked', this.checked); jQuery('#gform_export_check_all').html(this.checked ? '<strong><?php echo esc_js( __( 'Deselect All', 'gravityforms' ) ); ?></strong>' : '<strong><?php echo esc_js( __( 'Select All', 'gravityforms' ) ); ?></strong>'); \"> <label id='gform_export_check_all' for='select_all'><strong><?php esc_html_e( 'Select All', 'gravityforms' ) ?></strong></label></li>";
						for (var i = 0; i < aryFields.length; i++) {
							fieldList += "<li><input type='checkbox' id='export_field_" + i + "' name='export_field[]' value='" + aryFields[i][0] + "' class='gform_export_field'> <label for='export_field_" + i + "'>" + aryFields[i][1] + "</label></li>";
						}
						jQuery("#export_field_list").html(fieldList);
						jQuery("#export_date_start, #export_date_end").datepicker({dateFormat: 'yy-mm-dd', changeMonth: true, changeYear: true});

						jQuery("#export_field_container, #export_filter_container, #export_date_container, #export_submit_container").hide().show();

						gf_vars.filterAndAny = <?php echo json_encode( esc_html__( 'Export entries if {0} of the following match:', 'gravityforms' ) ); ?>;
						jQuery("#export_filters").gfFilterUI(filterSettings);
					}
					jQuery(document).ready(function () {
						jQuery("#gform_export").submit(function () {
							if (jQuery(".gform_export_field:checked").length == 0) {
								alert(<?php echo json_encode( __( 'Please select the fields to be exported', 'gravityforms' ) );  ?>);
								return false;
							}
						});
					});


				</script>

				<p class="textleft"><?php esc_html_e( 'Select a form below to export entries. Once you have selected a form you may select the fields you would like to export and then define optional filters for field values and the date range. When you click the download button below, Gravity Forms will create a XLSX file for you to save to your computer.', 'gravityforms' ); ?></p>
				<div class="hr-divider"></div>
				<form id="gform_export" method="post" style="margin-top:10px;">
					<?php echo wp_nonce_field( 'rg_start_export', 'rg_start_export_nonce' ); ?>
					<table class="form-table">
						<tr valign="top">

							<th scope="row">
								<label for="export_form"><?php esc_html_e( 'Select A Form', 'gravityforms' ); ?></label> <?php gform_tooltip( 'export_select_form' ) ?>
							</th>
							<td>

								<select id="export_form" name="export_form" onchange="SelectExportForm(jQuery(this).val());">
									<option value=""><?php esc_html_e( 'Select a form', 'gravityforms' ); ?></option>
									<?php
									$forms = RGFormsModel::get_forms( null, 'title' );
									foreach ( $forms as $form ) {
										?>
										<option value="<?php echo absint( $form->id ) ?>"><?php echo esc_html( $form->title ) ?></option>
									<?php
									}
									?>
								</select>

							</td>
						</tr>
						<tr id="export_field_container" valign="top" style="display: none;">
							<th scope="row">
								<label for="export_fields"><?php esc_html_e( 'Select Fields', 'gravityforms' ); ?></label> <?php gform_tooltip( 'export_select_fields' ) ?>
							</th>
							<td>
								<ul id="export_field_list">
								</ul>
							</td>
						</tr>
						<tr id="export_filter_container" valign="top" style="display: none;">
							<th scope="row">
								<label><?php esc_html_e( 'Conditional Logic', 'gravityforms' ); ?></label> <?php gform_tooltip( 'export_conditional_logic' ) ?>
							</th>
							<td>
								<div id="export_filters">
									<!--placeholder-->
								</div>

							</td>
						</tr>
						<tr id="export_date_container" valign="top" style="display: none;">
							<th scope="row">
								<label for="export_date"><?php esc_html_e( 'Select Date Range', 'gravityforms' ); ?></label> <?php gform_tooltip( 'export_date_range' ) ?>
							</th>
							<td>
								<div>
		                            <span style="width:150px; float:left; ">
		                                <input type="text" id="export_date_start" name="export_date_start" style="width:90%" />
		                                <strong><label for="export_date_start" style="display:block;"><?php esc_html_e( 'Start', 'gravityforms' ); ?></label></strong>
		                            </span>

		                            <span style="width:150px; float:left;">
		                                <input type="text" id="export_date_end" name="export_date_end" style="width:90%" />
		                                <strong><label for="export_date_end" style="display:block;"><?php esc_html_e( 'End', 'gravityforms' ); ?></label></strong>
		                            </span>

									<div style="clear: both;"></div>
									<?php esc_html_e( 'Date Range is optional, if no date range is selected all entries will be exported.', 'gravityforms' ); ?>
								</div>
							</td>
						</tr>
					</table>
					<ul>
						<li id="export_submit_container" style="display:none; clear:both;">
							<br /><br />
							<input type="submit" name="export_lead_excel" value="<?php esc_attr_e( 'Download Export File', 'gravityforms' ); ?>" class="button button-large button-primary" />
		                    <span id="please_wait_container" style="display:none; margin-left:15px;">
		                        <i class='gficon-gravityforms-spinner-icon gficon-spin'></i> <?php esc_html_e( 'Exporting entries. Please wait...', 'gravityforms' ); ?>
		                    </span>

							<iframe id="export_frame" width="1" height="1" src="about:blank"></iframe>
						</li>
					</ul>
				</form>
				<?php
				GFExport::page_footer();
		}
	}

	/**
	 * Register the stylesheets for the admin area.
	 *
	 * @since    1.0.0
	 */
	public function enqueue_styles() {

		/**
		 * This function is provided for demonstration purposes only.
		 *
		 * An instance of this class should be passed to the run() function
		 * defined in Gravity_Forms_Export_Entries_Excel_Loader as all of the hooks are defined
		 * in that particular class.
		 *
		 * The Gravity_Forms_Export_Entries_Excel_Loader will then create the relationship
		 * between the defined hooks and the functions defined in this
		 * class.
		 */

		wp_enqueue_style( $this->plugin_name, plugin_dir_url( __FILE__ ) . 'css/gravity-forms-export-entries-excel-admin.css', array(), $this->version, 'all' );

	}

	/**
	 * Register the JavaScript for the admin area.
	 *
	 * @since    1.0.0
	 */
	public function enqueue_scripts() {

		/**
		 * This function is provided for demonstration purposes only.
		 *
		 * An instance of this class should be passed to the run() function
		 * defined in Gravity_Forms_Export_Entries_Excel_Loader as all of the hooks are defined
		 * in that particular class.
		 *
		 * The Gravity_Forms_Export_Entries_Excel_Loader will then create the relationship
		 * between the defined hooks and the functions defined in this
		 * class.
		 */

		$scripts = array(
			'jquery-ui-datepicker',
			'gform_form_admin',
			'gform_field_filter',
			'sack',
		);

		if ( empty( $scripts ) ) {
			return;
		}

		foreach ( $scripts as $script ) {
			wp_enqueue_script( $script );
		}

	}

	public static function maybe_export_excel() {
		if ( isset( $_POST['export_lead_excel'] ) ) {
			check_admin_referer( 'rg_start_export', 'rg_start_export_nonce' );
			//see if any fields chosen
			if ( empty( $_POST['export_field'] ) ) {
				GFCommon::add_error_message( __( 'Please select the fields to be exported', 'gravityforms' ) );

				return;
			}
			$form_id = $_POST['export_form'];
			$form    = RGFormsModel::get_form_meta( $form_id );

			self::start_export_excel( $form );
			die();
		} else if ( isset( $_POST['export_forms'] ) ) {
			check_admin_referer( 'gf_export_forms', 'gf_export_forms_nonce' );
			$selected_forms = rgpost( 'gf_form_id' );
			if ( empty( $selected_forms ) ) {
				GFCommon::add_error_message( __( 'Please select the forms to be exported', 'gravityforms' ) );
				return;
			}

			$forms = RGFormsModel::get_form_meta_by_id( $selected_forms );

			// clean up a bit before exporting
			foreach ( $forms as &$form ) {

				foreach ( $form['fields'] as &$field ) {
					$inputType = RGFormsModel::get_input_type( $field );

					if ( isset( $field->pageNumber ) ) {
						unset( $field->pageNumber );
					}

					if ( $inputType != 'address' ) {
						unset( $field->addressType );
					}

					if ( $inputType != 'date' ) {
						unset( $field->calendarIconType );
						unset( $field->dateType );
					}

					if ( $inputType != 'creditcard' ) {
						unset( $field->creditCards );
					}

					if ( $field->type == $field->inputType ) {
						unset( $field->inputType );
					}

					// convert associative array to indexed
					if ( isset( $form['confirmations'] ) ) {
						$form['confirmations'] = array_values( $form['confirmations'] );
					}

					if ( isset( $form['notifications'] ) ) {
						$form['notifications'] = array_values( $form['notifications'] );
					}
				}
								
				/**
				 * Allows you to filter and modify the Export Form
				 *
				 * @param array $form Assign which Gravity Form to change the export form for
				 */
				$form = gf_apply_filters( array( 'gform_export_form', $form['id'] ), $form );

			}

			$forms['version'] = GFForms::$version;

			$forms_json = json_encode( $forms );

			$filename = 'gravityforms-export-' . date( 'Y-m-d' ) . '.json';
			header( 'Content-Description: File Transfer' );
			header( "Content-Disposition: attachment; filename=$filename" );
			header( 'Content-Type: application/json; charset=' . get_option( 'blog_charset' ), true );
			echo $forms_json;
			die();
		}
	}

	public static function start_export_excel( $form ) {

		$form_id = $form['id'];
		$fields  = $_POST['export_field'];

		$start_date = empty( $_POST['export_date_start'] ) ? '' : GFExport::get_gmt_date( $_POST['export_date_start'] . ' 00:00:00' );
		$end_date   = empty( $_POST['export_date_end'] ) ? '' : GFExport::get_gmt_date( $_POST['export_date_end'] . ' 23:59:59' );

		$search_criteria['status']        = 'active';
		$search_criteria['field_filters'] = GFCommon::get_field_filters_from_post( $form );
		if ( ! empty( $start_date ) ) {
			$search_criteria['start_date'] = $start_date;
		}

		if ( ! empty( $end_date ) ) {
			$search_criteria['end_date'] = $end_date;
		}

		$sorting = array( 'key' => 'date_created', 'direction' => 'DESC', 'type' => 'info' );

		GFCommon::log_debug( "GFExport::start_export(): Start date: {$start_date}" );
		GFCommon::log_debug( "GFExport::start_export(): End date: {$end_date}" );

		$form = GFExport::add_default_export_fields( $form );

		$entry_count = GFAPI::count_entries( $form_id, $search_criteria );

		$page_size = 100;
		$offset    = 0;

		//Adding BOM marker for UTF-8
		$lines = chr( 239 ) . chr( 187 ) . chr( 191 );

		// set the separater
		$separator = ";";

		$field_rows = GFExport::get_field_row_count( $form, $fields, $entry_count );

		//writing header
		$headers = array();
		foreach ( $fields as $field_id ) {
			$field = RGFormsModel::get_field( $form, $field_id );
			$label = gf_apply_filters( array( 'gform_entries_field_header_pre_export', $form_id, $field_id ), GFCommon::get_label( $field, $field_id ), $form, $field );
			$value = str_replace( '"', '""', $label );

			GFCommon::log_debug( "GFExport::start_export(): Header for field ID {$field_id}: {$value}" );

			if ( strpos( $value, '=' ) === 0 ) {
				// Prevent Excel formulas
				$value = "'" . $value;
			}

			$headers[ $field_id ] = $value;

			$subrow_count = isset( $field_rows[ $field_id ] ) ? intval( $field_rows[ $field_id ] ) : 0;
			if ( $subrow_count == 0 ) {
				$lines .= '"' . $value . '"' . $separator;
			} else {
				for ( $i = 1; $i <= $subrow_count; $i ++ ) {
					$lines .= '"' . $value . ' ' . $i . '"' . $separator;
				}
			}

			GFCommon::log_debug( "GFExport::start_export(): Lines: {$lines}" );
		}
		$lines = substr( $lines, 0, strlen( $lines ) - 1 ) . "\n";

		//paging through results for memory issues
		while ( $entry_count > 0 ) {

			$paging = array(
				'offset'    => $offset,
				'page_size' => $page_size,
			);
			$leads  = GFAPI::get_entries( $form_id, $search_criteria, $sorting, $paging );

			$leads = gf_apply_filters( array( 'gform_leads_before_export', $form_id ), $leads, $form, $paging );

			foreach ( $leads as $lead ) {
				foreach ( $fields as $field_id ) {
					switch ( $field_id ) {
						case 'date_created' :
							$lead_gmt_time   = mysql2date( 'G', $lead['date_created'] );
							$lead_local_time = GFCommon::get_local_timestamp( $lead_gmt_time );
							$value           = date_i18n( 'Y-m-d H:i:s', $lead_local_time, true );
							break;
						default :
							$field = RGFormsModel::get_field( $form, $field_id );

							$value = is_object( $field ) ? $field->get_value_export( $lead, $field_id, false, true ) : rgar( $lead, $field_id );
							$value = apply_filters( 'gform_export_field_value', $value, $form_id, $field_id, $lead );

							GFCommon::log_debug( "GFExport::start_export(): Value for field ID {$field_id}: {$value}" );
							break;
					}

					if ( isset( $field_rows[ $field_id ] ) ) {
						$list = empty( $value ) ? array() : unserialize( $value );

						foreach ( $list as $row ) {
							$row_values = array_values( $row );
							$row_str    = implode( '|', $row_values );

							if ( strpos( $row_str, '=' ) === 0 ) {
								// Prevent Excel formulas
								$row_str = "'" . $row_str;
							}

							$lines .= '"' . str_replace( '"', '""', $row_str ) . '"' . $separator;
						}

						//filling missing subrow columns (if any)
						$missing_count = intval( $field_rows[ $field_id ] ) - count( $list );
						for ( $i = 0; $i < $missing_count; $i ++ ) {
							$lines .= '""' . $separator;
						}
					} else {
						$value = maybe_unserialize( $value );
						if ( is_array( $value ) ) {
							$value = implode( '|', $value );
						}

						if ( strpos( $value, '=' ) === 0 ) {
							// Prevent Excel formulas
							$value = "'" . $value;
						}

						$lines .= '"' . str_replace( '"', '""', $value ) . '"' . $separator;
					}
				}
				$lines = substr( $lines, 0, strlen( $lines ) - 1 );

				GFCommon::log_debug( "GFExport::start_export(): Lines: {$lines}" );

				$lines .= "\n";
			}

			$offset += $page_size;
			$entry_count -= $page_size;

			if ( ! seems_utf8( $lines ) ) {
				$lines = utf8_encode( $lines );
			}
			$lines = apply_filters( 'gform_export_lines', $lines );
			//echo $lines;
			$Data = str_getcsv($lines, "\n"); //parse the rows 
			foreach($Data as &$Row) $Row = str_getcsv($Row, ";"); //parse the items in rows
			$Data[0][0] = str_replace('"', '', $Data[0][0]);
			// Create new PHPExcel object
			$objPHPExcel = new PHPExcel();
			// Add some data
			$objPHPExcel->getActiveSheet()->fromArray($Data, null, 'A1');
			$objPHPExcel->getActiveSheet()->getStyle('A1:PIG1')->getFont()->setBold(true);
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$objPHPExcel->setActiveSheetIndex(0);
			$filename = sanitize_title_with_dashes( $form['title'] ) . '-' . gmdate( 'Y-m-d', GFCommon::get_local_timestamp( time() ) ) . '.xlsx';
			header('Content-Type: application/vnd.ms-excel');
			header( "Content-Disposition: attachment; filename=$filename" );
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			$objWriter->save('php://output');
			exit();
			$lines = '';
		}

		/**
		 * Fires after exporting all the entries in form
		 *
		 * @param array $form The Form object to get the entries from
		 * @param string $start_date The start date for when the export of entries should take place
		 * @param string $end_date The end date for when the export of entries should stop
		 * @param array $fields The specified fields where the entries should be exported from
		 */
		do_action( 'gform_post_export_entries', $form, $start_date, $end_date, $fields );

	}

}
