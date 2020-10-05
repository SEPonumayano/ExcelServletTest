<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>
<body>
<form action="Exceledit" method="post">
<p>name<input type="text" name="name" value="${name}"></p>
<p>number<input type="text" name="number" value="${number}"></p>

<p>date<input type="date" name="date" value="${date}"></p>

<input type="hidden" name="value" value="${value}">

<p><input type="submit" value="送信"></p>

</form>
</body>
</html>