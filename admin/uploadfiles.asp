<%@ language="javascript"%>
<%
var self = Request.serverVariables("SCRIPT_NAME");
if (Request.serverVariables("REQUEST_METHOD")=="POST")
{
        var oo = new uploadFile();
        oo.path = "../upload";                        //存放路径,为空表示当前路径,默认为uploadFile
        oo.named = "file";                        //命名方式,date表示用日期来命名,file表示用文件名本身,默认为file
        oo.ext = "all";                                //允许上传的扩展名,all表示都允许,默认为all
        oo.over = true;                                //当存在相同文件名时是否覆盖,默认为false
        oo.size = 1*1024*1024*1024;                //最大字节数限制,默认为1G
        oo.upload();//

		//location.replace("'+self+'");
}

//ASP无组件上传类
function uploadFile()
{
    var bLen  = Request.totalBytes;
    var bText = Request.binaryRead(bLen);
    var oo = Server.createObject("ADODB.Stream");
    oo.mode = 3;
        this.path = "uploadFile";
        this.named = "file";
        this.ext = "all";
        this.over = false;
        this.size = 1*1024*1024*1024;        //1GB

        //文件上传        
        this.upload = function ()
        {
                var o = this.getInfo();
                if (o.size>this.size)
                {
                        alert("文件过大,不能上传!");
                        return;                
                }
                var f = this.getFileName();
                var ext = f.replace(/^.+\./,"");
                if (this.ext!="all"&&!new RegExp(this.ext.replace(/,/g,"|"),"ig").test(ext))
                {
                        alert("目前暂不支持扩展名为 "+ext+" 的文件上传!");
                        return;
                }
                if (this.named=="date")
                {
                        f = new Date().toLocaleString().replace(/\D/g,"") + "." + ext;
                }

                oo.open();
                oo.type = 1;
                oo.write(o.bin);
                this.path = this.path.replace(/[^\/\\]$/,"$&/");
                var fso = Server.createObject("Scripting.FileSystemObject");
                if(this.path!=""&&!fso.folderExists(Server.mapPath(this.path)))
                {
                        fso.createFolder(Server.mapPath(this.path));
                }
                try
                {
                        oo.saveToFile(Server.mapPath(this.path+f),this.over?2:1);
                        alert("上传成功!");
						alert(f);
				 		Response.write('<script type="text/javascript">parent.myform.DefaultPicUrl.value="upload/'+f+'";</script>');
						
                }
                catch(e)
                {
                        alert("对不起,此文件已存在！");
                }
                oo.close();
                delete(oo);

        }

        //获取二进制和文件字节数
        this.getInfo = function ()
        {
                oo.open();
                oo.type=1;
                oo.write(bText);
                oo.position = 0;                                
                oo.type=2;
                oo.charset="unicode";
                var gbCode=escape(oo.readText()).replace(/%u(..)(..)/g,"%$2%$1");
                var sPos=gbCode.indexOf("%0D%0A%0D%0A")+12;
                var sLength=bLen-(gbCode.substring(0,gbCode.indexOf("%0D%0A")).length/3)-sPos/3-6;
                oo.close();
        
                oo.open();
                oo.type = 1;        
                oo.write(bText);
                oo.position=sPos/3;
                var bFile=oo.read(sLength);
                oo.close();
                
                return { bin:bFile, size:sLength };
        }

        //获取文件名        
        this.getFileName = function ()
        {
                oo.open();
                oo.type = 2;
                oo.writeText(bText);
                oo.position = 0;
                oo.charset = "gb2312";
                var fileName = oo.readText().match(/filename=\"(.+?)\"/i)[1].split("\\").slice(-1)[0];
                oo.close();
                return fileName;
        }
        
        function alert(msg)
        {
                Response.write('<script type="text/javascript">alert("'+msg+'");</script>');
        }
}
%>
<html>

<head>
  <title>ASP无组件上传类</title>
  <meta http-equiv="content-Type" content="text/html; charset=gb2312">
  
</head>
<body>

  <form action="<%=self%>" method="post" enctype="multipart/form-data" onSubmit="return (this.upFile.value!='');"> 
    <input type="file" name="upFile"/>
    <input type="submit" value="上传文件"/>
  </form>
 
</body>
</html>
