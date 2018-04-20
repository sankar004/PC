set JAVA_HOME=C:\Program Files (x86)\Java\jre7
set PATH=C:\Windows;C:\Windows\System32;C:\Program Files (x86)\Java\jre7\bin
set JARS="H:\Pdf_box\Apachi POI jar files\*;H:\Email.jar;.;"

set mydate=%date:~4,2%%date:~7,2%%date:~10,4%
echo "Launching Java Application..."
java  -cp %JARS% Compas_Email.PDA_Outlook

timeout /t -1
