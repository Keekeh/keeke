<HTML>
<HEAD>
<meta charset="utf-8" name="viewport" content="width=device-width" >

</HEAD>
<title>Página ASP</title>
<link href="style.css" rel="stylesheet" type="text/css">
<BODY>
<div class="fundo"><%
if Hour(Now) < 12 then %>
<h1 class="noos">Bom Dia</h1><BR>
<% elseif Hour(Now) < 18 then %>
<center><h1>Boa tarde!</h1></center>
<% elseif Hour(Now) < 21 then %>
<center><h1>Boa Noite</h1></center>
<% elseif Hour(Now) < 22 then %>
<CENTER><h1>Está tarde, vá dormir!</h1> </CENTER>
<% else %>
<center><h1>Ainda acordado maníaco??</h1></center>
<% end if %>
<CENTER>
  <h2>Testando componentes gráfico e disposição de tela via CSS!</h2>
</CENTER>
</div>
<center>
<div><img src="../imagens/loira.jpg">
<img src="../imagens/yes.jpg">
<img src= "../imagens/padre2.jpg"></div>
</center>
<nav class="but">
	<ul>
		<li><a href="../index.html">Index D&iacute;ogenes G&ouml;tz</a></li>
	</ul>
</nav>

</BODY>
</HTML>