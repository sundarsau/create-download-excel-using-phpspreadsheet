<?php
  $server="localhost";
  $userid="root";
  $pwd="";
  $dbname="test";
  $conn = new mysqli($server, $userid, $pwd, $dbname);
//Check connection
if ($conn->connect_error){
  die("Connection Error: " .$conn->connect_error);
} 
  
