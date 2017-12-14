var Word, Doc;

var T,t;
var T1,t1;
var str="Control time:\n";

Word=new ActiveXObject("Word.Application");
Word.Visible=true;
Doc=Word.Documents.Add();
Doc.Content.Bold=1;
Doc.Content.Font.Size=18;

var fso;
fso=new ActiveXObject("Scripting.FileSystemObject");


t=new Date();
T=t.getTime();
delete t;

do{
t1=new Date();
T1=t1.getTime();
if( (T1-T)%5000==0){
  str+=t1.getDate()+'.';
  str+=t1.getMonth()+'.';
  str+=t1.getYear()+'\t';
  str+=t1.getHours()+':';
  str+=t1.getMinutes()+':';  str+=t1.getSeconds()+'\n';
  if(fso.FileExists("C:\\dummy.ddd"))  str+="The file has appeared\n";
  else str+="The file is absent\n";

  Doc.Content.Text=str;
}
delete t1;
}while(T1-T<60000);

WSH.Quit();
