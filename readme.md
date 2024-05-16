# DinosOffice: Delphi Components for LibreOffice

### Important
----

In the last update, the DocVisible property was included, by default it is false, so when generating the spreadsheet it is not displayed on the screen,
If you want to see the spreadsheet being generated, set it to true;

----
Need install:
https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/

 - 1 - Open project "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\OpenOfficeComponent_install.dproj"
         For delphi 7 use the C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackag\OpenOfficeComponent_install_Delphi7.dpk
 - 2 - Clean
 - 3 - Build
 - 4 - install
 - 5 - tools -> options -> language -> Delphi -> Library -> Library/ADD:
        "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src" 
        For delphi 7 use the "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src_dx7"
 
 
| Project {versionOficial}   	| Version 	   |
|----------------------------	|------------ |
| 19.5                     	  | tested ✅  |

Tested Delphi version

| Version  	| Supported 	|
|----------	|-----------	|
| > 12.x   	|    ✅ 	    |
| > 11.x   	|    ✅    	|
| > 10.x   	|    ✅ 	    |
| Seattle  	|    ✅ 	    |
| XE 8     	|    ✅ 	    |
| Delphi 7 	|    ✅    	|

## For Unigui 
 You need add FDGUIxWaitCursor to your serverModule and in your serverModule checked the property AutoCoInitialize
 
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/d20893ff-e2c0-4e37-a823-33be3175091e)
 
| Version   	              | Supported 	|
|-------------------------	|-----------	|
| > 1.95.0.1584 with D12 	 |    ✅ 	   |

----
For basic read and write sheet usage, just use the component
To style the worksheet and documents, use the additional units: uOpenOfficeHelper.pas, uOpenOfficeCollors.pas 


This software is MIT open source!

Wiki: https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/wiki

## See more on Instagram https://www.instagram.com/dinosdev/
