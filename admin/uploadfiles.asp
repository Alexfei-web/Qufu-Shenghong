<%@ language="javascript"%>
<%
var self = Request.serverVariables("SCRIPT_NAME");
if (Request.serverVariables("REQUEST_METHOD")=="POST")
{
        var oo = new uploadFile();
        oo.path = "../upload";                        //���·��,Ϊ�ձ�ʾ��ǰ·��,Ĭ��ΪuploadFile
        oo.named = "file";                        //������ʽ,date��ʾ������������,file��ʾ���ļ�������,Ĭ��Ϊfile
        oo.ext = "all";                                //�����ϴ�����չ��,all��ʾ������,Ĭ��Ϊall
        oo.over = true;                                //��������ͬ�ļ���ʱ�Ƿ񸲸�,Ĭ��Ϊfalse
        oo.size = 1*1024*1024*1024;                //����ֽ�������,Ĭ��Ϊ1G
        oo.upload();//

		//location.replace("'+self+'");
}

//ASP������ϴ���
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

        //�ļ��ϴ�        
        this.upload = function ()
        {
                var o = this.getInfo();
                if (o.size>this.size)
                {
                        alert("�ļ�����,�����ϴ�!");
                        return;                
                }
                var f = this.getFileName();
                var ext = f.replace(/^.+\./,"");
                if (this.ext!="all"&&!new RegExp(this.ext.replace(/,/g,"|"),"ig").test(ext))
                {
                        alert("Ŀǰ�ݲ�֧����չ��Ϊ "+ext+" ���ļ��ϴ�!");
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
                        alert("�ϴ��ɹ�!");
						alert(f);
				 		Response.write('<script type="text/javascript">parent.myform.DefaultPicUrl.value="upload/'+f+'";</script>');
						
                }
                catch(e)
                {
                        alert("�Բ���,���ļ��Ѵ��ڣ�");
                }
                oo.close();
                delete(oo);

        }

        //��ȡ�����ƺ��ļ��ֽ���
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

        //��ȡ�ļ���        
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
  <title>ASP������ϴ���</title>
  <meta http-equiv="content-Type" content="text/html; charset=gb2312">
  
</head>
<body>

  <form action="<%=self%>" method="post" enctype="multipart/form-data" onSubmit="return (this.upFile.value!='');"> 
    <input type="file" name="upFile"/>
    <input type="submit" value="�ϴ��ļ�"/>
  </form>
 
</body>
</html>
