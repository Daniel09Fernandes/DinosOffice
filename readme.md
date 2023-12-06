# Fork changes

----

I did smalls changes to the samples projects structure, now you can test the samples project without install the packages.
You can use the this framework without install the package too, just add the src folder in the search path project.
Maybe in the future, i will add some fixes towards memory leaks and code improvements.

**Thanks a lot Daniel Fernandes.**

# DinosOffice: Delphi Components for LibreOffice

### Important

----

In the last update, the DocVisible property was included, by default it is false, so when generating the spreadsheet it is not displayed on the screen,
If you want to see the spreadsheet being generated, set it to true;

----
Necessario instalar o libreoffice : 
Need to install :

https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/

 - 1 - Open project "C:\yourLocal\Component_OpenOffice\OpenOfficeComponent_install.dproj" (for delphi 7 use the C:\yourLocal\Component_OpenOffice\OpenOfficeComponent_install_Delphi7.dpk)
 - 2 - Clean
 - 3 - Build
 - 4 - install
 - 5 - tools -> options -> language -> Delphi -> Library -> Library/ADD "C:\yourLocal\Component_OpenOffice\Src" (for delphi 7 use the "C:\yourLocal\Component_OpenOffice\Src_dx7")

 
 
| Project {versionOficial}   	| Version 	   |
|----------------------------	|------------ |
| 19.5                     	  | tested ✅  |

Tested Delphi version

| Version  	| Supported 	|
|----------	|-----------	|
| > 12.x   	|         ✅ 	|
| > 11.x   	|         ✅ 	|
| > 10.x   	|         ✅ 	|
| Seattle  	|         ✅ 	|
| XE 8     	|         ✅ 	|
| Delphi 7 	|         ✅ 	|

Para uso básico de leitura e gravação de plhanilha, use apenas o componente

For basic read and write sheet usage, just use the component


Para estilizar a planilha e documentos, use as units adicionais: uOpenOfficeHelper.pas, uOpenOfficeCollors.pas 

To style the worksheet and documents, use the additional units: uOpenOfficeHelper.pas, uOpenOfficeCollors.pas 


This software is MIT open source!

Wiki: https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/wiki
