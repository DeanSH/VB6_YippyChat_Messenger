<?php
        session_start();      
// Settings
        $thename = $_GET["yname"];
        $save_path = '/home4/deano/etc/uploads/'; //getcwd() . '/uploads/';                            // The path were we will save the file (getcwd() may not be reliable and should be tested in your environment)
        $upload_name = 'filedata';                                                      // change this accordingly
        $max_file_size_in_bytes = 5242880; //2147483647;                           // 5mb - 2GB in bytes
        $whitelist = array('jpg', 'png', 'gif', 'jpeg', 'bmp', 'JPG', 'JPEG');        // Allowed file extensions
        $backlist = array('php', 'php3', 'php4', 'phtml','exe'); // Restrict file extensions
        $valid_chars_regex = 'A-Za-z0-9_-\s ';// Characters allowed in the file name (in a Regular Expression format)
         
// Other variables     
        $MAX_FILENAME_LENGTH = 260;
        $file_name = '';
        $file_extension = '';
        $uploadErrors = array(
        0=>'There is no error, the file uploaded with success',
        1=>'The uploaded file exceeds the upload_max_filesize directive in php.ini',
        2=>'The uploaded file exceeds the MAX_FILE_SIZE directive that was specified in the HTML form',
        3=>'The uploaded file was only partially uploaded',
        4=>'No file was uploaded',
        6=>'Missing a temporary folder'
        );
 
// Validate the upload
        //$tmpfilenme = $_FILES[$upload_name]['tmp_name'];
        if (!isset($_FILES[$upload_name]))
                HandleError('No upload found in \$_FILES for ' . $upload_name);
        else if (isset($_FILES[$upload_name]['error']) && $_FILES[$upload_name]['error'] != 0)
                HandleError($uploadErrors[$_FILES[$upload_name]['error']]);
        else if (!isset($_FILES[$upload_name]['tmp_name']) || !@is_uploaded_file($_FILES[$upload_name]['tmp_name']))
                HandleError('Upload failed is_uploaded_file test.');
        else if (!isset($_FILES[$upload_name]['name']))
                HandleError('File has no name.');
 
// Validate the file size (Warning: the largest files supported by this code is 2GB)
        $file_size = @filesize($_FILES[$upload_name]['tmp_name']);
        if (!$file_size || $file_size > $max_file_size_in_bytes)
                HandleError('File exceeds the maximum allowed size of 5MB');
         
        if ($file_size <= 0)
                HandleError('File size outside allowed lower bound');
// Validate its a MIME Images (Take note that not all MIME is the same across different browser, especially when its zip file)
        if(!eregi('image/', $_FILES[$upload_name]['type']))
                HandleError('Please upload a valid file!');
 
// Validate that it is an image
        $imageinfo = getimagesize($_FILES[$upload_name]['tmp_name']);
        if($imageinfo['mime'] != 'image/gif' && $imageinfo['mime'] != 'image/jpeg' && $imageinfo['mime'] != 'image/png' && isset($imageinfo))
                HandleError('Sorry, we only accept PNG , GIF and JPEG images');
 
// Validate file name (for our purposes we'll just remove invalid characters)
        $file_name = $_FILES[$upload_name]['name'];
        if (strlen($file_name) == 0 || strlen($file_name) > $MAX_FILENAME_LENGTH)
                HandleError('Invalid file name');
 
// Validate that we will over-write an existing file
        if (file_exists($save_path . $file_name))
                unlink ($save_path . $file_name);

// Validate file extension
         foreach ($backlist as $item) {
         if(preg_match("/$item$/i", $file_name)) {
         HandleError('Invalid file extension ' . $file_name);
         }
         }
        //if(!in_array(end(explode('.', $file_name)), $whitelist))
               // HandleError('Invalid file extension' . $file_name);
        //if(in_array(end(explode('.', $file_name)), $backlist))
               // HandleError('Invalid file extension2' . $file_name);
// Rename the file to be saved 
        $temp2 = explode(".", $_FILES[$upload_name]["name"]);
        $extension2 = end($temp2);
        $file_name = 'YacYo'.$thename.$extension2; //md5($file_name. time());
         
// Verify! Upload the file
        if (!@move_uploaded_file($_FILES[$upload_name]['tmp_name'], $save_path.$file_name)) {
                HandleError('File could not be saved.');
        }
        //unlink ($tmpfilenme);
        //echo 'Upload Successful ' . $thename;
        //exit(0);
        resize($thename, $save_path.'YacYo'.$thename, $save_path.$file_name);
        
/* Handles the error output. */
function HandleError($message) {
        //unlink ($tmpfilenme);
        echo $message;
        exit(0);
}

function resize($namey, $targetFile, $originalFile) {

    $info = getimagesize($originalFile);
    $mime = $info['mime'];

    switch ($mime) {
            case 'image/jpeg':
                    $image_create_func = 'imagecreatefromjpeg';
                    $image_save_func = 'imagejpeg';
                    $new_image_ext = 'jpg';
                    break;

            case 'image/png':
                    $image_create_func = 'imagecreatefrompng';
                    $image_save_func = 'imagepng';
                    $new_image_ext = 'png';
                    break;

            case 'image/gif':
                    $image_create_func = 'imagecreatefromgif';
                    $image_save_func = 'imagegif';
                    $new_image_ext = 'gif';
                    break;

            default: 
                    //throw Exception('Unknown image type.');
                    HandleError('Unknown image type for Resize.');
    }
    
    $img = $image_create_func($originalFile);
    list($width, $height) = getimagesize($originalFile);

    //$newHeight = $newWidth; //1280 x 800
    
    $thumb_width = 100;
    $thumb_height = 100;
    
    $original_aspect = $width / $height;
    $thumb_aspect = $thumb_width / $thumb_height;

if ( $original_aspect >= $thumb_aspect )
{
   // If image is wider than thumbnail (in aspect ratio sense)
   $new_height = $thumb_height;
   $new_width = $width / ($height / $thumb_height);
}
else
{
   // If the thumbnail is wider than the image
   $new_width = $thumb_width;
   $new_height = $height / ($width / $thumb_width);
}
    
    $tmpH = 0 - ($new_height - $thumb_height) / 2;
    $tmpW = 0 - ($new_width - $thumb_width) / 2;
    
    $tmp = imagecreatetruecolor($thumb_width, $thumb_height);
    imagecopyresampled($tmp, $img, $tmpW, $tmpH, 0, 0, $new_width, $new_height, $width, $height);

    if (file_exists($originalFile)) {
            unlink($originalFile);
    }
    if (file_exists("$targetFile.jpg")) {
            unlink("$targetFile.jpg");
    }
    //$image_save_func($tmp, $originalFile);
    imagejpeg($tmp, "$targetFile.jpg", 100);
        echo 'Upload Successful ' . $namey . ' tmpW ' . $tmpW . ' tmpH ' . $tmpH . ' new_width ' . $new_width . ' new_height ' . $new_height . ' width ' . $width . ' height ' . $height;
        exit(0);
}
?>