<?php
require_once('../../config.php');
require_once('customform.php');

global $CFG, $PAGE, $USER;
//require_login();

$PAGE->set_context(context_system::instance());
$PAGE->set_pagelayout('standard');
$PAGE->set_title('Office file reader');
$PAGE->set_heading('Office file reader');
$PAGE->set_url(new moodle_url('/local/officehelper/index.php'));

echo $OUTPUT->header();

$mform = new customform();
$toform = array();
$mform->set_data($toform);
if ($mform->is_cancelled()) {
    redirect(new moodle_url('/my'), 'Redirecting', 0);
} elseif ($mform->no_submit_button_pressed()) {
    $fs = get_file_storage();
    $usercontext = context_user::instance($USER->id);
    $draftitemid = file_get_submitted_draft_itemid('files');

    if (optional_param('read', 0, PARAM_RAW)) {
        $files = $fs->get_area_files($usercontext->id, 'user', 'draft',
            $draftitemid, 'filename', false);
        $mform->output($files);
    }
    $mform->display();
} else {
    $mform->display();
}

echo $OUTPUT->footer();
?>