acad = actxserver('AutoCAD.Application');
App = acad.invoke('IAcadApplication');
Doc = App.ActiveDocument;
MS = Doc.ModelSpace;
L = MS.Item(0); %handle to a line object

path = 'C:\Users\Arbi\Documents\frst.dxf'

set(acad,'visible',1);
MyDrawing  = acad.Application.Documents.Open(path)
MyDrawing.SendCommand("_Circle 2,2,0 4 ")
acad.Application.Documents.invoke.('_Circle 2,2,0 4 ')
MyDrawing = acad.ActiveDocument()
ThisDrawing = acad.Application

ThisDrawing.SendCommand "-dataextraction" & vbCr & "C:\ORTODRAIN UK\calcs.dxe" & vbCr

H=actxserver('AutoCAD.Application');
set(H,'visible',1);
H2=H.ActiveDocument;
H3=H2.ModelSpace;
AddCircle(H3,[0,0,0],100);

App.RunMacro("-dataextraction")

dxe_path = 'C:\Users\Arbi\Documents\first_data_extract.dxe'

Doc.SendCommand("_Circle 2,2,0 4 ")

Doc.SendCommand("-dataextraction C:\Users\Arbi\Documents\first_data_extract.dxe yes")

invoke(Doc,'PostCommand','_Circle vbCr 4,4,0 vbCr 4 vbCr') %create circle

invoke(Doc,'SendCommand', '-dataextraction C:\Users\Arbi\Documents\third_data_extract.dxe' & vbCr & 'Yes')

invoke(Doc,'PostCommand', 'Yes')

vbsFilePath = 'C:\Users\Arbi\Documents\writeTextFile.vbs';
command = ['Call runScript("',vbsFilePath,'")'];
invoke(acad,'processCommand',command);

invoke(Doc,'processCommand',command);

Dim count As Integer
count = invoke(MS, 'Count')

MS.Count
MS.Item('Arc')