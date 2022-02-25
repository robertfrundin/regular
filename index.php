<?php


require 'vendor/autoload.php';



$source = __DIR__."/streshnevo.docx";

$phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
$body="";
$sections=$phpWord->getSections();
foreach ($sections as $section) {
    $arrays = $section->getElements();

    foreach ($arrays as $el) {

        if (get_class($el) === 'PhpOffice\PhpWord\Element\TextRun') {
            foreach ($el->getElements() as $text) {


                if (get_class($text)==='PhpOffice\PhpWord\Element\TextBreak'){
                    $body.=' ';

                }
                else {
                    $body .= ' ';
                    $body .= $text->getText() . ' ';
                }



            }

        }
        elseif (get_class($el) === 'PhpOffice\PhpWord\Element\TextBreak'){
            $body.=' ';
        }
        elseif(get_class($el)==='PhpOffice\PhpWord\Element\Table'){
            $body .= ' ';

            $rows = $el->getRows();

            foreach($rows as $row) {
                $body .= ' ';

                $cells = $row->getCells();
                foreach($cells as $cell) {
                    $body .= ' ';
                    $celements = $cell->getElements();
                    foreach($celements as $celem) {
                        if(get_class($celem) === 'PhpOffice\PhpWord\Element\Text') {
                            $body .= $celem->getText();
                        }

                        else if(get_class($celem) === 'PhpOffice\PhpWord\Element\TextRun') {
                            foreach($celem->getElements() as $text) {
                                if (get_class($text)==='PhpOffice\PhpWord\Element\TextBreak'){
                                    $body.=' ';
                                }
                                else {
                                    $body .= $text->getText();
                                }

                            }
                        }
                    }
                    $body .= ' ';
                }

                $body .= ' ';
            }

            $body .= ' ';
        }
    }
    $body.=' ';
}
echo $body;

 preg_match_all('/(?:[1-5] кк)\s+(?:\d+(?:[.,]?\d+)?) млн/', $body, $rooms);

echo '<pre>' . print_r($rooms, 1) . '</pre>';

preg_match_all('/Срок\sсдачи.*[0-9]{1}\sкв\s[0-9]{4}/', $body, $deadlines);

echo '<pre>' . print_r($deadlines, 1) . '</pre>';

preg_match_all('/(без отделки)|(отделка)/', $body, $finishing);

echo '<pre>' . print_r($finishing, 1) . '</pre>';


?>



























