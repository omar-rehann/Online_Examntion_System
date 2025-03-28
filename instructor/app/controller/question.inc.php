
<?php
session_start();
include_once 'autoloader.inc.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;




if (isset($_GET['uploadImage'])) {
  $up = uploadFile($_FILES['file']['tmp_name']);
  echo '../style/images/uploads/' . $up . '.jpg';
}elseif (isset($_GET['deleteImage'])) {
    deleteImage($_POST['src']);
    echo 'success';
}elseif (isset($_GET['deleteAnswer'])){
  $q = new question;
  if(is_numeric($_POST['ansID'])){
    $q->deleteAnswer($_POST['ansID']);
    echo 'success';
  }
}elseif (isset($_GET['addQuestion'])){
    $question = isset($_POST['questionText']) ? trim($_POST['questionText']) : null;
    $qtype = isset($_POST['qtype']) ? trim($_POST['qtype']) : null;
    $isTrue = isset($_POST['isTrue']) ? trim($_POST['isTrue']) : 0;
    $points = isset($_POST['points']) ? trim($_POST['points']) : 0;
    $difficulty = isset($_POST['difficulty']) ? trim($_POST['difficulty']) : 1;
    $course = $_POST['Course'];
    if($question == null){
      $_SESSION["error"][] = 'Question Can\' Be Empty';
      header('Location: ' . $_SERVER['HTTP_REFERER']);exit;
    }elseif($qtype == null){
      $_SESSION["error"][] = 'Question Type is not selected';
      header('Location: ' . $_SERVER['HTTP_REFERER']);exit;
    }

    $newQuestion = new question;
    $newQuestion->insertQuestion($question,$qtype,$course,$isTrue,$points,$difficulty);
    $_SESSION["info"][] = 'Question Successfully Added';
    if ($qtype == 0) {
        foreach ($_POST['MCQanswer'] as $key=>$qanswer) {
            $answer = !empty($qanswer['answertext']) ? trim($qanswer['answertext']) : null;
            $isCorrect = !empty($qanswer['isCorrect']) ? 1 : 0;
            if ($answer != null) {
                $newQuestion->insertAnswersToLast($answer, $isCorrect,null);
            }
        }
    } elseif ($qtype == 3) {
      foreach ($_POST['MSQanswer'] as $key=>$qanswer) {
          $answer = !empty($qanswer['answertext']) ? trim($qanswer['answertext']) : null;
          $isCorrect = !empty($qanswer['isCorrect']) ? 1 : 0;
          if ($answer != null) {
              $newQuestion->insertAnswersToLast($answer, $isCorrect,null);
          }
      }
    }elseif ($qtype == 2) {
        foreach ($_POST['Canswer'] as $key=>$canswer) {
            $answer = $canswer['answertext'];
            if ($answer != '') {
                $newQuestion->insertAnswersToLast($answer, 1, null);
            }
        }
    }elseif ($qtype == 4) {
        foreach ($_POST['match'] as $key=>$manswer) {
            $matchAnswer = $_POST['matchAnswer'][$key];
            $matchPoints = $_POST['matchPoints'][$key];
            $answer = $manswer;
            if ($manswer != '' and $matchAnswer != '') {
                $newQuestion->insertAnswersToLast($manswer, 1, $matchAnswer,$matchPoints);
            }
        }
    }
    header('Location: ../../?questions=add&topic=' . $course);exit;
} elseif (isset($_GET['deleteQuestion'])) {
    $qst = new question;
    $qst->setQuestionDelete($_GET['deleteQuestion']);
    header('Location: ../../?questions');
} elseif (isset($_GET['restoreQuestion'])) {
    $qst = new question;
    $qst->restoreQuestion($_GET['restoreQuestion']);
    header('Location: ' . $_SERVER['HTTP_REFERER']);
} elseif (isset($_GET['PDeleteQuestion'])) {
    $qst = new question;
    $qst->pDeleteQuestion($_GET['PDeleteQuestion']);
    header('Location: ' . $_SERVER['HTTP_REFERER']);

} elseif (isset($_GET['updateQuestion'])) {
    $id = isset($_POST['qid']) ? trim($_POST['qid']) : null;
    $question = isset($_POST['questionText']) ? trim($_POST['questionText']) : null;
    $qtype = isset($_POST['qtype']) ? trim($_POST['qtype']) : 0;
    $isTrue = isset($_POST['isTrue']) ? trim($_POST['isTrue']) : 0;
    $points = isset($_POST['points']) ? trim($_POST['points']) : 0;
    $difficulty = isset($_POST['difficulty']) ? trim($_POST['difficulty']) : 1;
    $course = $_POST['Course'];

    $newQuestion = new question;
    $newQuestion->updateQuestion($id,$question,$course,$points,$difficulty);
    $newQuestion->updateTF($id, $isTrue);

    if ($qtype == 0 || $qtype == 3) {
        foreach ($_POST['Qanswer'] as $key=>$qanswer) {
            $ansID = isset($qanswer['ansID']) ? trim($qanswer['ansID']) : null;
            $answer = !empty($qanswer['answertext']) ? trim($qanswer['answertext']) : null;
            $isCorrect = !empty($qanswer['isCorrect']) ? trim($qanswer['isCorrect']) : 0;
            if ($ansID == null) {
                if ($answer != null) {
                    $newQuestion->insertAnswers($id, $answer, $isCorrect);
                }
              } else {
                $newQuestion->updateAnswer($ansID, $answer, $isCorrect,null);
            }
        }
    } elseif ($qtype == 2) {
        foreach ($_POST['Canswer'] as $key=>$canswer) {
            $answer = $canswer['answertext'];
            if ($answer != '') {
                $newQuestion->insertAnswers($id,$answer,1);
            }
        }
    } elseif ($qtype == 4) {
      foreach ($_POST['match'] as $key=>$manswer) {
          $oldAns = isset($_POST['oldID'][$key]) ? $_POST['oldID'][$key] : null;
          $matchAnswer = $_POST['matchAnswer'][$key];
          $matchPoints = $_POST['matchPoints'][$key];
          if ($manswer != '' and $matchAnswer != '') {
            if($oldAns == null){
              $newQuestion->insertAnswers($id,$manswer,1,$matchAnswer,$matchPoints);
            }else{
              $newQuestion->updateAnswer($oldAns, $manswer, 1,$matchAnswer,$matchPoints);
            }
          }
      }
    }
    header('Location: ' . $_SERVER['HTTP_REFERER']);
}elseif (isset($_GET['duplicateQuestion']) and is_numeric($_GET['duplicateQuestion'])){
        $id = $_GET['duplicateQuestion'];
        $q = new question;
        $q->duplicateQuestion($id);
        $newID = $q->getLastQuestion()->id;
        header('Location:../../?questions=view&id='. $newID);
}else if (isset($_GET['export'])) {
    try {
        ob_end_clean(); // Clear any previous output

        // Check if the course name is provided
        if (!isset($_POST['course']) || empty($_POST['course'])) {
            throw new Exception("Course not specified.");
        }

        $course = $_POST['course'];
        $q = new question; 
        $questions = $q->getByCourse($course);

        if (empty($questions)) {
            throw new Exception("No questions found for this course.");
        }

        $qTypes = [
            0 => 'Multiple Choice', 1 => 'True/False', 2 => 'Complete',
            3 => 'Multiple Select', 4 => 'Matching', 5 => 'Essay'
        ];

        $data = [];
        foreach ($questions as $question) {
            $id = $question->id;
            $quest = strip_tags($question->question);
            $type = $question->type;
            $difficulty = $question->difficulty;
            $typetext = $qTypes[$type];
            $points = $question->points;
            $isTrue = $question->isTrue;

            $answers = $q->getQuestionAnswers($id);
            $ans1 = $answers[0]->answer ?? '';
            $ans2 = $answers[1]->answer ?? '';
            $ans3 = $answers[2]->answer ?? '';
            $ans4 = $answers[3]->answer ?? '';

            if ($type == 0 || $type == 3) { // Multiple Choice & Multiple Select
                $ans1 = ($answers[0]->isCorrect ?? false) ? "#!$ans1" : $ans1;
                $ans2 = ($answers[1]->isCorrect ?? false) ? "#!$ans2" : $ans2;
                $ans3 = ($answers[2]->isCorrect ?? false) ? "#!$ans3" : $ans3;
                $ans4 = ($answers[3]->isCorrect ?? false) ? "#!$ans4" : $ans4;
            } elseif ($type == 4) { // Matching
                $ans1 = isset($answers[0]) ? "{$answers[0]->answer} >> {$answers[0]->matchAnswer}" : '';
                $ans2 = isset($answers[1]) ? "{$answers[1]->answer} >> {$answers[1]->matchAnswer}" : '';
                $ans3 = isset($answers[2]) ? "{$answers[2]->answer} >> {$answers[2]->matchAnswer}" : '';
                $ans4 = isset($answers[3]) ? "{$answers[3]->answer} >> {$answers[3]->matchAnswer}" : '';
            } elseif ($type == 1) { // True/False
                $ans1 = ($isTrue == 1) ? 'True' : 'False';
                $ans2 = $ans3 = $ans4 = '';
            } elseif ($type == 5) { // Essay
                $ans1 = $ans2 = $ans3 = $ans4 = '';
            }

            $data[] = [$quest, $typetext, $points, $difficulty, $ans1, $ans2, $ans3, $ans4];
        }

        // Create a new Excel file
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Add headers
        $headers = ['Question', 'Question Type', 'Points', 'Difficulty', 'Answer 1', 'Answer 2', 'Answer 3', 'Answer 4'];
        $col = 'A';
        foreach ($headers as $header) {
            $sheet->setCellValue($col . '1', $header);
            $col++;
        }

        // Style headers
        $headerStyle = [
            'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
            'fill' => ['fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID, 'startColor' => ['rgb' => '4472C4']]
        ];
        $sheet->getStyle('A1:H1')->applyFromArray($headerStyle);

        // Add data
        $row = 2;
        foreach ($data as $dataRow) {
            $col = 'A';
            foreach ($dataRow as $cell) {
                $sheet->setCellValue($col . $row, $cell);
                $col++;
            }

            // Alternate row styling
            if ($row % 2 == 0) {
                $sheet->getStyle("A$row:H$row")->getFill()
                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                    ->getStartColor()->setRGB('E9E9E9');
            }
            $row++;
        }

        // Auto-size columns
        foreach (range('A', 'H') as $column) {
            $sheet->getColumnDimension($column)->setAutoSize(true);
        }

        // Prepare the file for download
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $course . '_Questions.xlsx"');
        header('Cache-Control: max-age=0');

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;

    } catch (Exception $e) {
        die("An error occurred: " . $e->getMessage());
    }
}else if (isset($_GET['import']) && !empty($_FILES)) {
    if (!isset($_FILES['excel']['tmp_name']) || empty($_FILES['excel']['tmp_name'])) {
        die("Error: No file uploaded!");
    }

    $excelFile = $_FILES['excel']['tmp_name'];

    if (!file_exists($excelFile)) {
        die("Error: File not found on the server!");
    }
    if (!is_readable($excelFile)) {
        die("Error: File exists but cannot be read!");
    }

    try {
        $reader = IOFactory::createReaderForFile($excelFile);
        $spreadsheet = $reader->load($excelFile);

        $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

        if (empty($sheetData)) {
            die("Error: Excel file is empty!");
        }

        require_once '../model/question.class.php';
        if (!class_exists('question')) {
            die("Error: Question class not found!");
        }

        $q = new question();
        $course = $_POST['course'] ?? null;

        if (!$course) {
            die("Error: Course not specified!");
        }

        $qTypes = ['Multiple Choise' => 0, 'True/False' => 1, 'Complete' => 2, 'Multiple Select' => 3, 'Matching' => 4, 'Essay' => 5];

        foreach ($sheetData as $rowIndex => $row) {
            if ($rowIndex < 14) continue;

            $questionText = trim($row['A'] ?? '');
            $questionTypeText = trim($row['B'] ?? '');
            $points = trim($row['C'] ?? 0);
            $difficulty = trim($row['D'] ?? 1);

            if (!$questionText || !isset($qTypes[$questionTypeText])) {
                continue;
            }

            $qtype = $qTypes[$questionTypeText];

            $q->insertQuestion($questionText, $qtype, $course, 0, $points, $difficulty);
            $lastInsertedId = $q->getLastQuestion()->id;

            if ($qtype == 0 || $qtype == 3) {
                for ($i = 'E'; $i <= 'H'; $i++) {
                    $answerText = trim($row[$i] ?? '');
                    $isCorrect = strpos($answerText, '#!') === 0 ? 1 : 0;
                    $answerText = str_replace('#!', '', $answerText);

                    if (!empty($answerText)) {
                        $q->insertAnswers($lastInsertedId, $answerText, $isCorrect);
                    }
                }
            } elseif ($qtype == 4) {
                for ($i = 'E'; $i <= 'H'; $i++) {
                    if (!empty($row[$i])) {
                        $matchParts = explode('>>', $row[$i]);
                        if (count($matchParts) == 2) {
                            $q->insertAnswers($lastInsertedId, trim($matchParts[0]), 1, trim($matchParts[1]));
                        }
                    }
                }
            } elseif ($qtype == 1) {
                $isTrue = (strtolower($row['E']) === 'true') ? 1 : 0;
                $q->updateTF($lastInsertedId, $isTrue);
            } elseif ($qtype == 2) {
                for ($i = 'E'; $i <= 'H'; $i++) {
                    $answerText = trim($row[$i] ?? '');
                    if (!empty($answerText)) {
                        $q->insertAnswers($lastInsertedId, $answerText, 1);
                    }
                }
            }
        }

    } catch (Exception $e) {
        die("Error reading the file: " . $e->getMessage());
    }

    header('Location: ../../?questions');
    exit;
}

else{
    http_response_code(404);
  }
  
?>