<?php

/**  Plugin Name: Import Excel2Sql
 * Description: A plugin that allows you to bulk import posts from excel to your wordpress website. Get started now!
 * Plugin URI: https://www.syrup.co.zw
 * Author: Nyasha Chawanda
 * Author URI: https://www.linkedin.com/in/nyasha-chawanda-742300a9/
 * Version: 1.0.1
 *
 * Text Domain: Import Excel2Sql
 *
 * Import Excel2Sql  is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * any later version.
 *
 * Import Excel2Sql is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 */

global $wpdb;

class Excel2Posts
{

	function __construct()
		{
		add_action('admin_menu', array($this, 'Excelposts_tool'));
		}    
	function Excelposts_tool(){//add the Excel2Posts function under tools tab
		add_management_page( 'Import Excel2Sql', 'Import Excel2Sql', 
		'manage_options', 'Import-Excel2Sql', array($this, 'Excel2Posts_form'), 0 );
		}
			
    function Excel2Posts_form()//the form to import the excel sheet and upload data to database
		{
        ?>
        <h1>Import Posts From Excel to Wordpress </h1><hr>
		
		<form class="" action="" method="post" enctype="multipart/form-data">

		
		<h3> Posts Status </h3>
		<p>Select the status which the post will be acquire after succesully transfered
			, <br><b>PUBLISH</b> means the post will automaticlly published and appear as posts on<br>your website post page.
			Then if you choose <br><b>DRAFT</b> they will appear as drafts waiting to be published manually later</p>
			<select name = "poststatus">
				<option value = "publish">Publish</option>
				<option value = "draft">Draft</option>
			</select><br><br><hr>


			<h3> Select the user to be atributed as the author of posts </h3>
			<select name = "postsauthor" required>
				<?php 
				$the_users = get_users();
					foreach ($the_users as $user){
						?>
					<option value = "<?php echo esc_attr($user->ID);?>"> 
					<?php echo  esc_attr( $user->display_name);?> </option>
				<?php }?>
			</select>



				
			</select><br><br><hr>
			<h3> Upload Excel file</h3>
			<p> To ensure you succesfully post without errors, make sure in your excel sheet
				<br>The first column is for <b>Post Title</b> and 
				<br>The second
				column is for <b> Post Content</b>. 
				<br>Dont add any other column content except these to columns. Supported excel exentions are .csv , .xls and .xlsx only
			</p>
			<input type="file" name="excel" accept =".csv, .xls, .xlsx" required value="">
			<button style = "background-color: orange; border-radius: 5px;padding: 5px;"type="submit" name="import">Start Importing</button>
			<br><br>
			
		</form>
        <?php
        
        //the import form excel

		if(isset($_POST["import"]))
			{
				$fileName = sanitize_text_field($_FILES["excel"]["name"]);
				$fileExtension = explode('.', $fileName);
				$fileExtension = strtolower(end($fileExtension));
				$newFileName = date("Y.m.d") . " - " . date("h.i.sa") . "." . $fileExtension;
				$targetDirectory = 'uploadedExcel/' . $newFileName;
				$uploadit = wp_upload_bits($newFileName, null, file_get_contents($_FILES['excel']['tmp_name']));
				$post_status = sanitize_text_field($_POST["poststatus"]);
				$post_author = sanitize_text_field($_POST["postsauthor"]);
				

			if ($uploadit)
				{// if the file have been succesfully uploaded to uploads folder

				require 'readexcel/excel_reader2.php';// import the module that read excel sheets
				require 'readexcel/SpreadsheetReader.php';
				//try reading the excel sheet and input the contents into the wordpress posts
				try{
					$reader = new SpreadsheetReader($uploadit['file']);// read the excel file
					foreach($reader as $key => $row){// iterate the data from excel sheet then post to wp_post table
						$posts_from_excel = array(
							'post_title'    =>$row[0],
							'post_content'  => $row[1],
							'post_status'   => $post_status, // status set on the form
							'post_type'   => 'post',
							'post_author'   => $post_author,
							//'post_category' => array( 8,39 )
						);
						wp_insert_post( $posts_from_excel );
						}
					echo wp_kses_post('<p style = "color: green;">Posts successfully transfered from excel to wordpress</p>');
					}
				catch (Exception $ex) {
					echo wp_kses_post('<p style = "color: red;">The transfer Failed. An error occured!</p>');
					}
				
				}
			}

    }

}
new Excel2Posts();// run the class 



