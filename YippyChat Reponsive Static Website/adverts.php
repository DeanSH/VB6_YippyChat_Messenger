<link rel="icon" href="/favicon.png" type="image/x-icon"/>
<link rel="shortcut icon" href="/favicon.png" type="image/x-icon"/>
<title>Adverts</title>
<script type="text/javascript">
var image1 = new Image()
var link1 = "http://www.yippychat.com/"
image1.src = "/yippylogo.png"
var image2 = new Image()
var link2 = "http://www.yippychat.com/"
image2.src = "/sponsors/sponsor1.jpg"
var image3 = new Image()
var link3 = "http://www.yippychat.com/"
image3.src = "/sponsors/sponsor2.jpg"
</script>
<center>
<a href="http://www.yippychat.com/" id="links"><img src="/yippylogo.png" width="275" height="60" name="slide" style="border-style: none"/></a>
<script type="text/javascript">
        var step=1;
        function slideit()
        {
            document.images.slide.src = eval("image"+step+".src");
            document.getElementById("links").href = eval("link"+step+"");
            if(step<3)
                step++;
            else
                step=1;
            setTimeout("slideit()",8000);
        }
        slideit();
</script>
</center>
