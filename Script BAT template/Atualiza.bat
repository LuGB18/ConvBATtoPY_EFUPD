color f4
ECHO *** Efetuando Limpeza da Base ***
cd ..
del *ACBrNFeServicos*.ini
cd Arquivos
md ..\Lixo
move ..\*.pdf ..\Lixo
move ..\*.csv ..\Lixo
move ..\*.xls ..\Lixo
move ..\*.rpf ..\Lixo
del *.rar
del *.zip
del "Plano de Contas Básico.doc"
del "Formatos dos Códigos.doc"
del "EFBackup1160010.exe"
del "EFBackup1160011.exe"
del IBPT.exe
del ..\Eficaz3*.exe
del ..\Eficaz - old*.exe
del ..\*old*.exe
del ..\Eficaz4*.exe
del ..\EFUpdate1*.exe
del ..\email.exe
del "..\Impressão*"
del ..\Schemas.exe
del ..\EFBackup1*.exe
del ..\dll.exe
del ..\Atualizacao0*.exe
del ..\Atualizacao0*.rar
del ..\EficazHelp0*.chm
del ..\Arquivos\Plano*.doc
del ..\Arquivos\Forma*.doc
del ..\EFUpdate\*exe
del ..\EFUpdate\*txt
del ..\EFUpdate\EficazHelp0*.chm
del ..\EFUpdate\EficazHelp0*.pdf
del ..\Utilitarios\Firebird-1*
del ..\Utilitarios\Firebird-2*
del ..\Utilitarios\Firebird-3*.
del ..\Utilitarios\Firebird-4*.
del ..\Utilitarios\InstaladorCadeiaV2*
del ..\Utilitarios\InstaladorCadeiaV5*
del ..\Utilitarios\Team*.exe
del ..\Utilitarios\IBExpert_2013*zip
del ..\Utilitarios\IBExpert_2013*rar
del ..\Utilitarios\IBExpert_2010*exe
del ..\Utilitarios\IBExpert_2010*rar
del ..\Utilitarios\EFBackup1160001.exe
del ..\Utilitarios\EFBackup1160002.exe
del ..\Utilitarios\EFBackup1160003.exe
del ..\Utilitarios\EFBackup1160004.exe
del ..\Utilitarios\EFBackup1160005.exe
del ..\Utilitarios\EFBackup1160006.exe
del ..\Utilitarios\EFBackup1160007.exe
del ..\Utilitarios\EFBackup1160008.exe
del ..\Utilitarios\EFBackup1160009.exe
del ..\Utilitarios\EFBackup1160010.exe
del ..\Utilitarios\Suporte*
del ..\Utilitarios\Show*
del ..\Utilitarios\Atualizacao011*.
del ..\Utilitarios\Atualizacao012*.
del ..\Utilitarios\Atualizacao013*.
del ..\Utilitarios\Atualizacao014*.
del ..\Utilitarios\Atualizacao015*
del ..\Utilitarios\Atualizacao016*
del ..\Utilitarios\Atualizacao017*
del ..\Utilitarios\Atualizacao018*
del ..\Utilitarios\Atualizacao019*
del ..\Utilitarios\Atualizacao020*
del ..\EFupdate2*
del ..\EFupdate1*
del ..\EFupdateB*
del ..\Utilitarios\Atualizacao.exe
del ..\Utilitarios\AA*.exe
del ..\Utilitarios\AA*.log
del ..\Utilitarios\Ammy*.exe
del ..\Utilitarios\Ammy*.log
rd ..\dll
rd ..\Capicom
rd ..\IBPT

ECHO *** Atualizando Capicom ***
rd Capicom
Capicom.exe

if EXIST %windir%\SysWOW64 goto Win64
:Win32
ECHO *** Copiando as DLLs ***
if NOT EXIST %windir%\System32\capicom.dll copy Capicom\capicom.dll %windir%\System32
if NOT EXIST %windir%\System32\msxml5.dll  copy Capicom\msxml5.dll  %windir%\System32
if NOT EXIST %windir%\System32\msxml5r.dll copy Capicom\msxml5r.dll %windir%\System32
ECHO *** Registrando as DLLs ***
regsvr32 %windir%\System32\capicom.dll /s
regsvr32 %windir%\System32\msxml5.dll /s
goto end
:Win64
ECHO *** Copiando as DLLs x64 ***
if NOT EXIST %windir%\SysWOW64\capicom.dll copy Capicom\capicom.dll %windir%\SysWOW64
if NOT EXIST %windir%\SysWOW64\msxml5.dll  copy Capicom\msxml5.dll  %windir%\SysWOW64
if NOT EXIST %windir%\SysWOW64\msxml5r.dll copy Capicom\msxml5r.dll %windir%\SysWOW64
ECHO *** Registrando as DLLs x64 ***
regsvr32 %windir%\SysWOW64\capicom.dll /s
regsvr32 %windir%\SysWOW64\msxml5.dll /s
goto end
:end

ECHO *** Atualizando as DLLs ***
rd dll /s /q
dll.exe
copy dll\*.* ..\ /y
rd dll /s /q

ECHO *** Atualizando as DLLs do EFServer ***
rd dll /s /q
dll.exe
copy dll\*.* ..\..\EFServer /y
rd dll /s /q

ECHO *** Atualizando as DLLs do EFServer ***
rd dll /s /q
dll.exe
copy dll\*.* ..\EFServer\ /y
rd dll /s /q

ECHO *** Atualizando Schemas ***
rd Schemas /s /q
rd ..\Schemas /s /q
Schemas.exe
move /y Schemas ..\

ECHO *** Atualizando Schemas-NFSe***
rd Schemas-NFSe /s /q
rd ..\Schemas-NFSe /s /q
Schemas-NFSe.exe
move /y Schemas-NFSe ..\

ECHO *** Atualizando Schemas-CTe ***
rd Schemas-CTe /s /q
rd ..\Schemas-CTe /s /q
Schemas-CTe.exe
move /y Schemas-CTe ..\

ECHO *** Atualizando Schemas-MDF-e ***
rd Schemas-MDFe /s /q
rd ..\Schemas-MDFe /s /q
Schemas-MDFe.exe
move /y Schemas-MDFe ..\

ECHO *** Atualizando Schemas-DFe ***
rd Schemas-DFe /s /q
rd ..\Schemas-DFe /s /q
Schemas-DFe.exe
move /y Schemas-DFe ..\

ECHO *** Limpando Logs ***
del ..\Atualizacao\*.sql

ECHO *** Atualizando Openssl ***
copy ..\Utilitarios\openssl.cnf ..\
copy ..\Utilitarios\openssl.exe ..\

ECHO *** Atualizando Help ***
copy ..\EFUpdate\EficazHelp.chm ..\

copy TrocaExe.bat ..\EFUpdate /y

del ..\EFUpdate\Servicos\*zip
del ..\EFUpdate\"Versao Beta"\*zip

ECHO *** Iniciando Backup em segundo plano ***
start "" /b cmd /c "limpabackup.bat"