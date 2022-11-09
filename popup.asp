<script language="javascript" type="text/javascript">
    function myPop() {
        this.square = null;
        this.overdiv = null;

        this.popOut = function(msgtxt) {
            //filter:alpha(opacity=25);-moz-opacity:.25;opacity:.25;
            this.overdiv = document.createElement("div");
            this.overdiv.className = "overdiv";

            this.square = document.createElement("div");
            this.square.className = "square";
            this.square.Code = this;
            var msg = document.createElement("div");
            msg.className = "msg";
            msg.innerHTML = msgtxt;
            this.square.appendChild(msg);
            var closebtn = document.createElement("button");
            closebtn.onclick = function() {
                this.parentNode.Code.popIn();
            }
            closebtn.innerHTML = "Biết rồi";
            this.square.appendChild(closebtn);

            document.body.appendChild(this.overdiv);
            document.body.appendChild(this.square);
        }
        this.popIn = function() {
            if (this.square != null) {
                document.body.removeChild(this.square);
                this.square = null;
            }
            if (this.overdiv != null) {
                document.body.removeChild(this.overdiv);
                this.overdiv = null;
            }

        }
    }

</script>

<style type="text/css">
 div.overdiv { filter: alpha(opacity=75);
                      -moz-opacity: .75;
                      opacity: .75;
                      background-color: #c0c0c0;
                      position: absolute;
                      top: 0px;
                      left: 0px;
                      width: 100%; height: 100%; }

        div.square { position: absolute;
                     top: 200px;
                     left: 200px;
                     background-color: Menu;
                     border: #f9f9f9;
                     height: 200px;
                     width: 300px; 
                     text-align:center;}
        div.square div.msg { color: #3e6bc2;
                             font-size: 15px;
                             padding: 15px; }
</style>

<html> 
  <head>
    <script type="text/javascript" src="NAME OF THE PAGE!.js"></script>
    <style>
        div.overdiv { filter: alpha(opacity=75);
                      -moz-opacity: .75;
                      opacity: .75;
                      background-color: #c0c0c0;
                      position: absolute;
                      top: 0px;
                      left: 0px;
                      width: 100%; height: 100%; }

        div.square { position: absolute;
                     top: 200px;
                     left: 200px;
                     background-color: Menu;
                     border: #f9f9f9;
                     height: 200px;
                     width: 300px; }
        div.square div.msg { color: #3e6bc2;
                             font-size: 15px;
                             padding: 15px; }
    </style>
  </head>
  <body>
    <div style="background-color: red; width: 200px; height: 300px;
                padding: 20px; margin: 20px;"></div>

    <script type="text/javascript">
        var pop = new myPop();
        pop.popOut("<h2> Thông báo </h2> Tiệc sơ kết sẽ diễn ra vào lúc 18h00 ngày 25/06/2010 tại nhà hàng Sáu Linh số 223 Nguyễn Thái Sơn, phường 7, Gò Vấp");
    </script>
  </body>
</html>
