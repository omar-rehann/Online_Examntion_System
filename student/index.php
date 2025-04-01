<?php
session_start();
date_default_timezone_set('Africa/Cairo');
define('NotDirectAccess', true);
require_once 'app/model/student.class.php';
require_once 'app/model/test.class.php';
require_once 'app/model/group.class.php';

if (isset($_SESSION['student'])) {
    $_student = new student;
    if (!$_student->checkSession($_SESSION['student']->id)) {
        unset($_SESSION['student']);
        header("Location: /");
        exit;
    }
    
    if (isset($_GET['results']))
        require_once 'app/view/results.php';
    elseif (isset($_GET['groups']))
        require_once 'app/view/groups.php';
    elseif (isset($_GET['profile']))
        require_once 'app/view/profile.php';
    elseif (isset($_GET['logout'])) {
        unset($_SESSION['student']);
        header("Location: /");
    }
    else
        require_once 'app/view/tests.php';
} else {
    if (isset($_GET['reset']))
        require_once 'app/view/reset.php';
    else
        require_once 'app/view/login.php';
}