<?php


require 'vendor/autoload.php';





function getDataVoxhall($filename)
{
    $source = __DIR__ . "/" . $filename;

    $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
    $body = "";
    $sections = $phpWord->getSections();


    foreach ($sections as $section) {
        $arrays = $section->getElements();

        foreach ($arrays as $el) {

            if (get_class($el) === 'PhpOffice\PhpWord\Element\TextRun') {
                foreach ($el->getElements() as $text) {


                    if (get_class($text) === 'PhpOffice\PhpWord\Element\TextBreak') {
                        $body .= ' ';

                    } else {
                        $body .= ' ';
                        $body .= $text->getText() . ' ';
                    }


                }

            } elseif (get_class($el) === 'PhpOffice\PhpWord\Element\TextBreak') {
                $body .= ' ';
            } elseif (get_class($el) === 'PhpOffice\PhpWord\Element\Table') {
                $body .= ' ';

                $rows = $el->getRows();

                foreach ($rows as $row) {
                    $body .= ' ';

                    $cells = $row->getCells();
                    foreach ($cells as $cell) {
                        $body .= ' ';
                        $celements = $cell->getElements();
                        foreach ($celements as $celem) {
                            if (get_class($celem) === 'PhpOffice\PhpWord\Element\Text') {
                                $body .= $celem->getText();
                            } else if (get_class($celem) === 'PhpOffice\PhpWord\Element\TextRun') {
                                foreach ($celem->getElements() as $text) {
                                    if (get_class($text) === 'PhpOffice\PhpWord\Element\TextBreak') {
                                        $body .= ' ';
                                    } else {
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
        $body .= ' ';
    }

//?????????????????? ???????????? ????????????
    function RoomPriceSplit($str)
    {
        $str = trim($str);
        $mas = ['', '', ''];
        preg_match_all('/[0-9]{2},[0-9]\s??????|[0-9]{1,2}\s??????/', $str, $res);
        preg_match_all('/([0-9]{1}\s+????)|(????????????|????????????|????????????|????????????)/', $str, $res2);
        preg_match_all('/???? [0-9]{2},[0-9]\s??2/', $str, $res3);
        $mas[1] = $res[0][0];
        $mas[0] = $res2[0][0];
        $mas[2] = $res3[0][0];


        return $mas;

    }


    preg_match_all('/(?:[1-5] ????)\s+(?:\d+(?:[.,]?\d+)?) ?????? ???? [0-9]{2},[0-9]\s??2/', $body, $rooms);
    preg_match_all('/(:???\s*????????????????|??????\s*??????????????)/', $body, $finishing);
    $flag = false;
    if (count($finishing[0]) < count($rooms[0])) {
        $flag = true;
    }
    echo '<table border=1px>';
    echo '<thead>
    <tr>
      <th style="
    padding: 3px;
">?????? ????????????
      </th>
      <th style="
    padding: 3px;
">????????
      </th>
      <th style="
    padding: 3px;
">??????????????
      </th>
     <th style="
    padding: 3px;
">??????????????
      </th>
    </tr>
  </thead>
  <tbody>';
    foreach ($rooms[0] as $el) {

        echo '<tr>
<td style="
    padding: 3px;
">';
        echo RoomPriceSplit($el)[0];
        echo '</td>';
        echo ' <td style="
    padding: 3px;
">';
        echo RoomPriceSplit($el)[1];
        echo '</td>';
        echo ' <td style="
    padding: 3px;
">';
        echo RoomPriceSplit($el)[2];
        echo '</td>';
        if ($flag === true) {
            echo ' <td style="
    padding: 3px;
">';
            echo $finishing[0][0];
            echo '</td>';
        }

        echo '</tr>';

    }


    echo '</tbody>';
    echo '</table>';
}
getDataVoxhall('voxhall.docx');

?>



























