<?php
class dbh {
  private $host = "localhost";
  private $port = "3306"; // أو 3307 حسب إعدادات XAMPP
  private $username = "root";
  private $pwd = ""; // تأكد من أنها فارغة إذا لم يكن هناك كلمة مرور
  private $dbName = "final";

  public function connect(){
    try {
      $conn = 'mysql:host=' . $this->host . ';port=' . $this->port . ';dbname=' . $this->dbName . ';charset=utf8mb4';
      $pdo = new PDO($conn, $this->username, $this->pwd);
      $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
      $pdo->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
      return $pdo;
    } catch (PDOException $e) {
      die("Database connection failed: " . $e->getMessage());
    }
  }
}
