<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lismount Tex Login</title>
<link rel="stylesheet" href="style.css">
<script>
function login(){
    var user=document.getElementById("username").value;
    var pass=document.getElementById("password").value;

    if(user=="admin" && pass=="1234"){
        window.location="index.html";
    } else {
        var alertBox = document.getElementById("alert");
        alertBox.style.display="block";
        alertBox.innerHTML="Wrong Login Details";
    }
}
</script>
</head>
<body>
<div class="login-box">
<h2>LISMOUNT TEX LOGIN</h2>
<input type="text" id="username" placeholder="Username">
<input type="password" id="password" placeholder="Password">
<button onclick="login()">Login</button>
<div class="alert" id="alert"></div>
</div>
</body>
</html>
