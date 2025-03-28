<?php
if (!defined('NotDirectAccess')) {
    die('Direct Access is not allowed to this page');
}
require_once 'header.php';
require_once 'navbar.php';

$res = new test;

if (isset($_GET['id'])) {
    if ($_GET['id'] == 'Last') {
        $result = $res->getFinishedResult();
    } else {
        $result = $res->getMyResult($_GET['id']);
    }

    if (empty($result)) {
        header('Location: ?results');
        exit;
    }

    $percent = ($result->TestDegree > 0) ? round(($result->FinalGrade / $result->TestDegree) * 100) : 0;
    $answers = $res->getResultAnswers($result->id);
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <style>
        .no-print { display: block; }
        @media print {
            .no-print { display: none; }
        }
    </style>
</head>
<body class="bg-light">
    <div class="container py-4">
        <div class="col-md-12 mx-auto">
            <?php if ($result->releaseResult == 0) : ?>
                <div class="alert alert-warning text-center" role="alert">
                    <i class="fas fa-exclamation-triangle mr-2"></i> This result has not been released yet
                </div>
            <?php else : ?>
                <div class="card shadow-sm">
                    <div class="card-header bg-white d-flex justify-content-between align-items-center">
                        <div>
                            <i class="fas fa-clipboard-check mr-2"></i>
                            <span class="font-weight-bold">Test Results - <?php echo strip_tags(htmlspecialchars_decode($result->testName)); ?></span>
                        </div>
                        <button type="button" class="btn btn-sm btn-primary no-print" onclick="window.print();">
                            <i class="fas fa-print mr-1"></i> Print
                        </button>
                    </div>
                    
                    <div class="card-body">
                        <?php if ($result->Questions == 0) : ?>
                            <div class="alert alert-danger text-center" role="alert">
                                <i class="fas fa-times-circle mr-2"></i> This test was not submitted
                            </div>
                        <?php else : ?>
                            <?php if ($percent >= $result->passPercent) : ?>
                                <div class="alert alert-success text-center py-3">
                                    <h4 class="alert-heading mb-3"><i class="fas fa-check-circle mr-2"></i>Congratulations! You passed the test</h4>
                                    <p class="h4 mb-0">Your score: <?php echo $percent; ?>%</p>
                                </div>
                            <?php else : ?>
                                <div class="alert alert-danger text-center py-3">
                                    <h4 class="alert-heading mb-3"><i class="fas fa-exclamation-circle mr-2"></i>Sorry, you didn't pass the test</h4>
                                    <p class="h4 mb-0">Your score: <?php echo $percent; ?>%</p>
                                </div>
                            <?php endif; ?>
                        <?php endif; ?>

                        <!-- Test Information Section -->
                        <div class="card mb-4 shadow-sm">
                            <div class="card-header bg-white">
                                <h5 class="mb-0"><i class="fas fa-info-circle mr-2"></i>Test Information</h5>
                            </div>
                            <div class="card-body">
                                <div class="row">
                                    <div class="col-md-6">
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Test Name:</span>
                                            <span class="float-right"><?php echo strip_tags(htmlspecialchars_decode($result->testName)); ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Instructor:</span>
                                            <span class="float-right"><?php echo strip_tags(htmlspecialchars_decode($result->Instructor)); ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Date:</span>
                                            <span class="float-right"><?php echo date('m/d/Y', strtotime($result->startTime)); ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Duration:</span>
                                            <span class="float-right"><?php echo $result->resultDuration; ?> minutes</span>
                                        </div>
                                    </div>
                                    <div class="col-md-6">
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Start Time:</span>
                                            <span class="float-right"><?php echo date('h:i A', strtotime($result->startTime)); ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">End Time:</span>
                                            <span class="float-right"><?php echo date('h:i A', strtotime($result->endTime)); ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Questions:</span>
                                            <span class="badge badge-primary float-right"><?php echo $result->Questions; ?></span>
                                        </div>
                                        <div class="mb-3 pb-2 border-bottom">
                                            <span class="font-weight-bold text-muted">Total Points:</span>
                                            <span class="badge badge-primary float-right"><?php echo $result->TestDegree; ?></span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <?php if ($result->Questions > 0) : ?>
                        <!-- Performance Analysis Section -->
                        <div class="card mb-4 shadow-sm">
                            <div class="card-header bg-white">
                                <h5 class="mb-0"><i class="fas fa-chart-pie mr-2"></i>Performance Analysis</h5>
                            </div>
                            <div class="card-body">
                                <div class="row">
                                    <div class="col-md-6 mb-4">
                                        <div class="card h-100">
                                            <div class="card-body text-center">
                                                <h5 class="card-title">Percentage Score</h5>
                                                <div class="chart-container">
                                                    <canvas id="percentChart"></canvas>
                                                </div>
                                                <h3 class="mt-3 <?php echo ($percent >= $result->passPercent) ? 'text-success' : 'text-danger'; ?>">
                                                    <?php echo $percent; ?>%
                                                </h3>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-md-6 mb-4">
                                        <div class="card h-100">
                                            <div class="card-body text-center">
                                                <h5 class="card-title">Final Grade</h5>
                                                <div class="chart-container">
                                                    <canvas id="gradeChart"></canvas>
                                                </div>
                                                <h3 class="mt-3 <?php echo ($percent >= $result->passPercent) ? 'text-success' : 'text-danger'; ?>">
                                                    <?php echo $result->FinalGrade; ?>/<?php echo $result->TestDegree; ?>
                                                </h3>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="row mt-2">
                                    <div class="col-md-6">
                                        <div class="d-flex justify-content-between align-items-center p-3 bg-light rounded">
                                            <span class="font-weight-bold">Correct Answers:</span>
                                            <span class="badge badge-success badge-pill py-2 px-3"><?php echo $result->RightQuestions; ?></span>
                                        </div>
                                    </div>
                                    <div class="col-md-6">
                                        <div class="d-flex justify-content-between align-items-center p-3 bg-light rounded">
                                            <span class="font-weight-bold">Wrong Answers:</span>
                                            <span class="badge badge-danger badge-pill py-2 px-3"><?php echo $result->WrongQuestions; ?></span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <?php endif; ?>
                        
                        <?php if ($result->Questions > 0) : ?>
                        <!-- Questions Section -->
                        <div class="card shadow-sm">
                            <div class="card-header bg-white">
                                <h5 class="mb-0"><i class="fas fa-question-circle mr-2"></i>Questions and Answers</h5>
                            </div>
                            <div class="card-body">
                                <?php
                                $viewAnswers = $res->canViewResults($result->id);
                                if (($viewAnswers == 0) || (($viewAnswers == 1) && (strtotime($result->testEnd) < strtotime("now")))) {
                                    $types = [0 => 'Multiple Choice', 1 => 'True/False', 2 => 'Fill in Blank', 3 => 'Multiple Select', 4 => 'Matching', 5 => 'Essay'];
                                    $I = 0;
                                    foreach ($answers as $answer) {
                                        $I++;
                                ?>
                                        <div class="card mb-4 border">
                                            <div class="card-header bg-light d-flex justify-content-between align-items-center">
                                                <div>
                                                    <span class="badge badge-secondary mr-2"><?php echo $I; ?></span>
                                                    <span><?php echo $types[$answer->type]; ?></span>
                                                </div>
                                                <div>
                                                    <?php if ($answer->points == 0) : ?>
                                                        <span class="badge badge-danger">0 points</span>
                                                    <?php elseif ($answer->points > 0) : ?>
                                                        <span class="badge badge-success">+<?php echo $answer->points; ?> points</span>
                                                    <?php else : ?>
                                                        <span class="badge badge-warning">Not graded yet</span>
                                                    <?php endif; ?>
                                                </div>
                                            </div>
                                            <div class="card-body">
                                                <h6 class="card-title"><?php echo strip_tags(htmlspecialchars_decode($answer->question)); ?></h6>
                                                <hr>
                                                
                                                <?php if ($answer->type == 4) : ?>
                                                    <!-- Matching Type -->
                                                    <div class="row mt-3">
                                                        <div class="col-md-6">
                                                            <h6>Your Answer:</h6>
                                                            <?php $givenAnswers = $res->getGivenAnswers($result->id, $answer->id); ?>
                                                            <table class="table table-sm table-bordered">
                                                                <thead class="thead-light">
                                                                    <tr>
                                                                        <th>Answer</th>
                                                                        <th>Match</th>
                                                                        <th>Points</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody>
                                                                    <?php foreach ($givenAnswers as $ans) : ?>
                                                                        <tr class="<?php echo ($ans->isCorrect) ? 'table-success' : 'table-danger'; ?>">
                                                                            <td><?php echo strip_tags(htmlspecialchars_decode($ans->answer)); ?></td>
                                                                            <td><?php echo strip_tags(htmlspecialchars_decode($ans->textAnswer)); ?></td>
                                                                            <td>+<?php echo $ans->points; ?></td>
                                                                        </tr>
                                                                    <?php endforeach; ?>
                                                                </tbody>
                                                            </table>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <h6>Correct Answer:</h6>
                                                            <?php $correctAnswers = $res->getCorrectAnswers($answer->id); ?>
                                                            <table class="table table-sm table-bordered">
                                                                <thead class="thead-light">
                                                                    <tr>
                                                                        <th>Answer</th>
                                                                        <th>Correct Match</th>
                                                                    </tr>
                                                                </thead>
                                                                <tbody>
                                                                    <?php foreach ($correctAnswers as $ans) : ?>
                                                                        <tr class="table-success">
                                                                            <td><?php echo strip_tags(htmlspecialchars_decode($ans->answer)); ?></td>
                                                                            <td><?php echo strip_tags(htmlspecialchars_decode($ans->matchAnswer)); ?></td>
                                                                        </tr>
                                                                    <?php endforeach; ?>
                                                                </tbody>
                                                            </table>
                                                        </div>
                                                    </div>
                                                
                                                <?php elseif ($answer->type == 0 || $answer->type == 3) : ?>
                                                    <!-- Multiple Choice / Multiple Select -->
                                                    <div class="mt-3">
                                                        <h6>Your Answer:</h6>
                                                        <?php
                                                        $givenAnswers = $res->getGivenAnswers($result->id, $answer->id);
                                                        $correctAnswers = $res->getCorrectAnswers($answer->id);
                                                        ?>
                                                        <div class="row">
                                                            <?php foreach ($givenAnswers as $ans) : ?>
                                                                <div class="col-md-6 mb-2">
                                                                    <div class="p-2 rounded <?php echo ($ans->isCorrect) ? 'bg-success text-white' : 'bg-danger text-white'; ?>">
                                                                        <?php echo strip_tags(htmlspecialchars_decode($ans->answer)); ?>
                                                                    </div>
                                                                </div>
                                                            <?php endforeach; ?>
                                                        </div>
                                                        
                                                        <?php $rand = rand(10, 100); ?>
                                                        <button class="btn btn-sm btn-outline-primary mt-2" type="button" data-toggle="collapse" data-target="#mcqCorrect<?php echo $rand; ?>">
                                                            <i class="fas fa-eye mr-1"></i>View Correct Answer
                                                        </button>
                                                        <div class="collapse mt-2" id="mcqCorrect<?php echo $rand; ?>">
                                                            <div class="card card-body">
                                                                <h6>Correct Answer:</h6>
                                                                <div class="row">
                                                                    <?php foreach ($correctAnswers as $ans) : ?>
                                                                        <div class="col-md-6 mb-2">
                                                                            <div class="p-2 rounded bg-success text-white">
                                                                                <?php echo strip_tags(htmlspecialchars_decode($ans->answer)); ?>
                                                                            </div>
                                                                        </div>
                                                                    <?php endforeach; ?>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                
                                                <?php elseif ($answer->type == 1) : ?>
                                                    <!-- True/False -->
                                                    <?php $givenAnswers = $res->getGivenAnswers($result->id, $answer->id); ?>
                                                    <div class="row">
                                                        <div class="col-md-6">
                                                            <h6>Your Answer:</h6>
                                                            <span class="badge <?php echo ($givenAnswers[0]->isTrue == $answer->isTrue) ? 'badge-success' : 'badge-danger'; ?>">
                                                                <?php echo ($givenAnswers[0]->isTrue) ? 'True' : 'False'; ?>
                                                            </span>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <h6>Correct Answer:</h6>
                                                            <span class="badge badge-success">
                                                                <?php echo ($answer->isTrue) ? 'True' : 'False'; ?>
                                                            </span>
                                                        </div>
                                                    </div>
                                                
                                                <?php elseif ($answer->type == 2) : ?>
                                                    <!-- Fill in Blank -->
                                                    <?php
                                                    $givenAnswers = $res->getGivenAnswers($result->id, $answer->id);
                                                    $correctAnswers = $res->getCorrectAnswers($answer->id);
                                                    ?>
                                                    <div class="row">
                                                        <div class="col-md-6">
                                                            <h6>Your Answer:</h6>
                                                            <div class="p-2 bg-light rounded">
                                                                <?php echo strip_tags(htmlspecialchars_decode($givenAnswers[0]->textAnswer)); ?>
                                                            </div>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <h6>Correct Answer:</h6>
                                                            <div>
                                                                <?php foreach ($correctAnswers as $a) : ?>
                                                                    <span class="badge badge-success mr-1 mb-1"><?php echo strip_tags(htmlspecialchars_decode($a->answer)); ?></span>
                                                                <?php endforeach; ?>
                                                            </div>
                                                        </div>
                                                    </div>
                                                
                                                <?php elseif ($answer->type == 5) : ?>
                                                    <!-- Essay -->
                                                    <?php $givenAnswers = $res->getGivenAnswers($result->id, $answer->id); ?>
                                                    <div>
                                                        <h6>Your Answer:</h6>
                                                        <div class="p-2 bg-light rounded">
                                                            <?php echo nl2br(strip_tags(htmlspecialchars_decode($givenAnswers[0]->textAnswer))); ?>
                                                        </div>
                                                    </div>
                                                <?php endif; ?>
                                            </div>
                                        </div>
                                <?php
                                    }
                                } elseif ($viewAnswers == 2) {
                                    echo '<div class="alert alert-warning text-center" role="alert">
                                        <i class="fas fa-ban mr-2"></i> You cannot view answers for this test
                                    </div>';
                                } else {
                                    echo '<div class="alert alert-info text-center" role="alert">
                                        <i class="fas fa-clock mr-2"></i> Your answers will be available on ' . date('m/d/Y h:i A', strtotime($result->testEnd)) . '
                                    </div>';
                                }
                                ?>
                            </div>
                        </div>
                        <?php endif; ?>
                    </div>
                </div>
            <?php endif; ?>
        </div>
    </div>

    <?php if ($result->Questions > 0) : ?>
    <script>
        // Percentage Chart
        $(document).ready(function() {
            const percentCtx = document.getElementById('percentChart').getContext('2d');
            const percentChart = new Chart(percentCtx, {
                type: 'doughnut',
                data: {
                    labels: ['Correct', 'Incorrect'],
                    datasets: [{
                        data: [<?php echo $percent; ?>, <?php echo max(0, 100 - $percent); ?>],
                        backgroundColor: ['#28a745', '#dc3545'],
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutoutPercentage: 70,
                    legend: { display: false },
                    tooltips: { enabled: false },
                    animation: {
                        animateScale: true,
                        animateRotate: true
                    }
                }
            });

            // Grade Chart
            const gradeCtx = document.getElementById('gradeChart').getContext('2d');
            const gradeChart = new Chart(gradeCtx, {
                type: 'doughnut',
                data: {
                    labels: ['Your Score', 'Remaining'],
                    datasets: [{
                        data: [<?php echo $result->FinalGrade; ?>, <?php echo max(0, $result->TestDegree - $result->FinalGrade); ?>],
                        backgroundColor: [
                            '<?php echo ($percent >= $result->passPercent) ? "#28a745" : "#dc3545"; ?>',
                            '#e9ecef'
                        ],
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutoutPercentage: 70,
                    legend: { display: false },
                    tooltips: { enabled: false },
                    animation: {
                        animateScale: true,
                        animateRotate: true
                    }
                }
            });
        });
    </script>
    <?php endif; ?>
</body>
</html>
<?php
} else {
    // All results listing page
    $results = $res->getMyResults();
?>
    <div class="container mt-4">
        <div class="row">
            <div class="col-md-12">
                <div class="card shadow-sm">
                    <div class="card-header bg-white">
                        <strong class="card-title"><i class="fas fa-list-alt mr-2"></i>All Results</strong>
                    </div>
                    <div class="card-body">
                        <table class="table table-striped table-bordered resultTable">
                            <thead>
                                <tr>
                                    <th>Test</th>
                                    <th>Date</th>
                                    <th>Result</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                <?php foreach ($results as $result) : 
                                    $percent = ($result->TestDegree > 0) ? round(($result->FinalGrade / $result->TestDegree) * 100) : 0;
                                ?>
                                    <tr>
                                        <td><?php echo strip_tags(htmlspecialchars_decode($result->testName)); ?></td>
                                        <td><?php echo date('m/d/Y', strtotime($result->endTime)); ?></td>
                                        <td>
                                            <?php if ($result->releaseResult == 1) : ?>
                                                <span class="badge badge-<?php echo ($percent >= 50) ? "success" : "danger"; ?>">
                                                    <?php echo $percent; ?>%
                                                </span>
                                            <?php else : ?>
                                                <span class="badge badge-warning">Under Review</span>
                                            <?php endif; ?>
                                        </td>
                                        <td>
                                            <?php if ($result->releaseResult == 1) : ?>
                                                <a href="?results&id=<?php echo $result->id; ?>" class="btn btn-sm btn-outline-primary">
                                                    <i class="fas fa-info-circle mr-1"></i>Details
                                                </a>
                                            <?php endif; ?>
                                        </td>
                                    </tr>
                                <?php endforeach; ?>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
<?php
}

define('ContainsDatatables', true);
require_once 'footer.php';
?>