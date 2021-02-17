<?php
require_once("$CFG->libdir/formslib.php");

use PhpOffice\PhpSpreadsheet\IOFactory as Excel;
use PhpOffice\PhpWord\IOFactory as Word;
use PhpOffice\PhpPresentation\IOFactory as PowerPoint;

class customform extends moodleform {
    public function definition() {
        $mform = $this->_form;

        $mform->addElement('filemanager', 'files', 'Files to read', null, array('subdirs' => false));
        $mform->setType('files', PARAM_FILE);

        $mform->registerNoSubmitButton('read');
        $group = array();
        $group[] = $mform->createElement('submit', 'read', 'Read');
        //$group[] = $mform->createElement('submit', 'reload', 'Reload');
        $group[] = $mform->createElement('cancel');
        $mform->addElement('group', 'buttons', '', $group, null, false);
    }

    public function output($files) {
        global $CFG;
        require_once("$CFG->dirroot/local/officehelper/vendor/autoload.php");
        $mform = $this->_form;

        $group = array();
        $group[] = $mform->createElement('html',
            html_writer::div($this->file_linker($this->generate_output_files($files))));
        $mform->addElement('group', 'result', 'Resulting files', $group, null, false);
    }

    /** @param stored_file[] $files */
    public function generate_output_files($files) {
        $dir = make_request_directory();
        $outputfilepaths = array();
        foreach ($files as $file) {
            $filepath = $dir.'\\'.$file->get_filename();
            $file->copy_content_to($filepath);
            if ($this->save_output_file($filepath)) {
                $outputfilepaths[$file->get_filename()] = $filepath;
            }
        }
        if (empty($outputfilepaths))
            return array();
        return $this->create_moodle_files($outputfilepaths);
    }

    private function save_output_file($filepath) {
        try {
            $extension = strtolower(pathinfo($filepath, PATHINFO_EXTENSION));
            if ($type = $this->classify_excel($extension)) {
                $reader = Excel::createReader($type);
                $reader->setIncludeCharts(true);
                $writer = Excel::createWriter($reader->load($filepath), $type);
                unlink($filepath);
                $writer->save($filepath);
            } elseif ($type = $this->classify_word($extension)) {
                $reader = Word::createReader($type);
                $writer = Word::createWriter($reader->load($filepath), $type);
                unlink($filepath);
                $writer->save($filepath);
            } elseif ($type = $this->classify_powerpoint($extension)) {
                $reader = PowerPoint::createReader($type);
                $writer = PowerPoint::createWriter($reader->load($filepath), $type);
                unlink($filepath);
                $writer->save($filepath);
            }
        } catch (Throwable $th) {
            return false;
        }
        return true;
    }

    private function create_moodle_files(array $files) {
        global $USER;

        $draftitemid = 0;
        file_prepare_draft_area($draftitemid, null, null, null, null);
        $fs = get_file_storage();
        $filerecord = new stdClass();
        $filerecord->contextid = context_user::instance($USER->id)->id;
        $filerecord->component = 'user';
        $filerecord->filearea = 'draft';
        $filerecord->itemid = $draftitemid;
        $filerecord->filepath = '/';

        $outputfiles = array();
        foreach ($files as $name => $path) {
            $filerecord->filename = $name;
            if (file_exists($path))
                $outputfiles[] = $fs->create_file_from_pathname($filerecord, $path);
        }
        return $outputfiles;
    }

    private function classify_excel($extension) {
        switch ($extension) {
            case 'xlsx': // Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm': // Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return 'Xlsx';
            case 'xls': // Excel (BIFF) Spreadsheet
            case 'xlt': // Excel (BIFF) Template
                return 'Xls';
            case 'ods': // Open/Libre Offic Calc
            case 'ots': // Open/Libre Offic Calc Template
                return 'Ods';
            case 'slk':
                return 'Slk';
            case 'xml': // Excel 2003 SpreadSheetML
                return 'Xml';
            case 'gnumeric':
                return 'Gnumeric';
            case 'htm':
            case 'html':
                return 'Html';
            default:
                return null;
        }
    }

    private function classify_word($extension) {
        switch ($extension) {
            case 'docx':
                return 'Word2007';
            case 'odf':
                return 'ODText';
            case 'rtf':
                return 'RTF';
            case 'htm':
            case 'html':
                return 'Html';
            default:
                return null;
        }
    }

    private function classify_powerpoint($extension) {
        switch ($extension) {
            case 'pptx':
                return 'PowerPoint2007';
            case 'ppt':
                return 'PowerPoint97';
            case 'odp':
                return 'ODPresentation';
            default:
                return null;
        }
    }

    function validation($data, $files) {
        return array();
    }

    function file_linker($files) {
        /** @var $OUTPUT core_renderer */
        global $OUTPUT;
        if (empty($files))
            return '';
        $output = array();
        foreach ($files as $file) {
            $output[] = html_writer::tag('p', html_writer::link($this->get_url($file),
                $OUTPUT->pix_icon(file_file_icon($file), get_mimetype_description($file),
                    'moodle', array('class' => 'icon')) . ' ' . s($file->get_filename())));
        }
        return implode($output);
    }

    /** @param stored_file $file */
    function get_url($file) {
        return moodle_url::make_draftfile_url($file->get_itemid(),
            $file->get_filepath(), $file->get_filename(), true);
    }
}