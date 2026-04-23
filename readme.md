# DinosOffice: Delphi Components for LibreOffice
<img width="256" height="256" alt="image" src="https://github.com/Daniel09Fernandes/DinosOffice/blob/versionOficial/img/logoCalc.png" />


### Important
----

the DocVisible property was included, by default it is false, so when generating the spreadsheet it is not displayed on the screen,
If you want to see the spreadsheet being generated, set it to true;

## Running in: 
### Vcl, Fmx(Win32/64), Unigui, Intraweb, MCPServer

![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/ddec6e5c-ea7e-4840-a080-facc7e2384bf)
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/ec6738a6-f775-4d99-996e-5a6fdc092a73) 
<img width="151" height="162" alt="Unigui" src="https://github.com/user-attachments/assets/28a0707a-1a3f-459b-a4aa-0c3be4e67e0f" /> &nbsp;
<img width="256" height="256" alt="image" src="https://github.com/user-attachments/assets/e913ded5-a2a9-461b-9977-0b99d1a42545" />
<img width="200" height="200" alt="image" src="https://github.com/user-attachments/assets/0291a693-c60d-4752-8930-980c1e9838ae" />

---
## About 🎯
 DinosOffice is an open source project for creating, customizing, and reading spreadsheets quickly and easily.
 
---- 
# Installation ⚙️

Need install:
https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/

Installation is done using the [`boss install`](https://github.com/HashLoad/boss) command:
``` sh
boss install github.com/Daniel09Fernandes/DinosOffice
```

Install via [`Embarcadero - GetIt`](https://getitnow.embarcadero.com/dinosoffice-libreoffice-for-delphi/)

OR manually:

 - 1 - Open project "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\OpenOfficeComponent_install.dproj"
         For delphi 7 use the C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackag\OpenOfficeComponent_install_Delphi7.dpk
 - 2 - Clean
 - 3 - Build
 - 4 - install
 - 5 - tools -> options -> language -> Delphi -> Library -> Library/ADD:
        "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src" 
        For delphi 7 use the "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src_dx7"
 
 
| Github Project {branch versionOficial}   	| Version - On Embarcadero GetIt 	| Version 	   |
|-------------------------------------------|---------------------------------|-----------  | 
| 20.4                     	                |20.3                             | Tested ✅   | 

Tested Delphi version

| Version  	| Supported 	|   Windows version Tested    |
|----------	|-----------	| --------------------------- | 
| Seattle  	|    ✅ 	    | Win7, Win 10, Win 11        |
| XE 8     	|    ✅ 	    | Win7, Win 10, Win 11        |
| Delphi 7 	|    ✅    	 | Win7, Win 10                |
| > 10.x   	|    ✅ 	    | Win 10                      |
| > 11.x   	|    ✅ 	    | Win 10, Win 11              |
| > 12.x   	|    ✅ 	    | Win 11                      |
| > 13.X   	|    ✅ 	    | Win 11                      |

| Frameworks Web   	       | Supported 	|
|-------------------------	|-----------	|
| Unigui 1.95.0.1584 	     |    ✅ 	   |
| Intraweb 14    	         |    ✅ 	   |

## For Unigui 
 You need add FDGUIxWaitCursor to your serverModule and in your serverModule checked the property AutoCoInitialize
 
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/d20893ff-e2c0-4e37-a823-33be3175091e)


## For Run on IIS 
Need give permission to users IIS

![image](https://github.com/user-attachments/assets/5b32db4d-b648-442d-ace4-08c6dce829dd)

![image](https://github.com/user-attachments/assets/edbdaec3-6d92-433d-9d35-8fa055bb1de1)

To Unigui on IIS, use this path to access your spreadsheet

![image](https://github.com/user-attachments/assets/50d70568-95d8-4e32-ad6c-a1829e71f055)



## For Intraweb
 You need add FDGUIxWaitCursor to your ServerController and in your ServerController change the property ComInitialize for ciMultiThreaded
 

 ![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/a164b806-ca33-4242-a183-1a62a6882e7b)

---
## Run MCPServer 
 In you MCP Client apont to DinosOfficeMCP.exe (use the STDIO protocol).

Claude IA exemple 

Access the configuration on developer and edit config

<img width="364" height="167" alt="image" src="https://github.com/user-attachments/assets/b0e1d02c-3806-49ec-81a2-b1bef521e93b" /><br>


<img width="945" height="685" alt="image" src="https://github.com/user-attachments/assets/5cdc6748-6437-4ed6-84ff-e8cfe8502de3" /><br>


In mcpServers node, attach DinosOfficeMCP.exe


<img width="1058" height="403" alt="image" src="https://github.com/user-attachments/assets/7c43bd89-1832-4964-8b5e-217574ae775a" />
<br>

There you go, now you can read and create spreadsheets with AI

----
For basic read and write sheet usage, just use the component 
To style the worksheet and documents, use the additional units: uOpenOfficeHelper.pas, uOpenOfficeCollors.pas 

Calc(Spreedsheet) Example
```pascal
var
   OpenOffice_calc1: TOpenOffice_calc;
begin
   OpenOffice_calc1 := TOpenOffice_calc.Create(Self);

  OpenOffice_calc1.DocVisible := False; //True to visible doc on screen
  OpenOffice_calc1.StartSheet
    .SetValue(1, 'A', 'STATUS')
       .SetBorder([bAll], opBrown)
       .changeJustify(fthRIGHT, ftvTOP)
       .SetBold(true)
       .changeFont('Arial', 12)
       .SetUnderline(true)
       .SetColor(opWhite, opMagenta)
    .setValue(1, 'B', 'VALOR')
       .changeJustify(fthRIGHT, ftvTOP)
       .SetBorder([bAll], opBrown)
       .SetBold(true)
       .ChangeFont('Arial', 12)
       .SetUnderline(true)
       .SetColor(opWhite, opMagenta)
    .setValue(2, 'B', 109, ftNumeric)
       .SetBorder([bAll], opBrown)
    .setValue(2, 'A', 'AGUA')
       .SetBorder([bAll], opBrown)
    .setValue(3, 'B', 105.55, ftNumeric)
       .SetBorder([bAll], opBrown)
    .setValue(3, 'A', 'LUZ')
       .SetBorder([bAll], opBrown)
    .setValue(4, 'B', 1005.22, ftNumeric)
    .setValue(4, 'A', 'ALUGUEL')
    .setValue(6, 'A', 'Total de linhas');

  OpenOffice_calc1.GetCountRow(lOutCountRow); 

  OpenOffice_calc1
    .setValue(6, 'B', lOutCountRow, ftNumeric)
    .setValue(7, 'A', 'Total de Colunas');

   OpenOffice_calc1.GetCountCell(lOutCountCell);

   OpenOffice_calc1
    .setValue(7, 'B', lOutCountCell, ftNumeric)
    .addNewSheet('A Receber', 1)
    .setValue(1, 'A', 'VALOR')
      .SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP)
      .setBold(true)
    .setValue(1, 'B', 'DESC')
      .SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP)
      .setBold(true)
      .changeFont('Arial', 12)
      .setUnderline(true)
      .setColor(opWhite, opCiano)
    .setValue(2, 'B', 200, ftNumeric)
    .setValue(2, 'A', 'Emprestimo')
    .setValue(3, 'B', 369.55, ftNumeric)
    .setValue(3, 'A', 'Dividendos')
    .setValue(4, 'B', 1585.22, ftNumeric)
    .setValue(4, 'A', 'ALUGUEL')
    .setValue(5, 'B', 1585.22, ftNumeric)
    .setValue(5, 'A', 'Renda extra')
    .setValue(6, 'B', 1585.22, ftNumeric)
    .setValue(6, 'A', 'ALUGUEL 2')
    .setValue(9, 'A', 'Total de linhas')
      .SetBold(True)
      .SetCellWidth(5000);

   OpenOffice_calc1.GetCountRow(lOutCountRow);

   OpenOffice_calc1
    .setValue(9, 'B', lOutCountRow, ftNumeric)
    .setValue(10, 'A', 'Total de Colunas')
    .SetBold(True);

   OpenOffice_calc1.GetCountCell(lOutCountCell);

   OpenOffice_calc1
    .setValue(10, 'B', lOutCountCell, ftNumeric)
    .setValue(8, 'A', 'Total')
      .setBold(true)
    .setFormula(8, 'B', '=B2+B3+B4+B5+B6')
      .setBold(true)
    .positionSheetByName(lSheetname);

  // Configure the chart settings
  SettingsChart.Height := 11000;
  SettingsChart.Width := 22000;
  SettingsChart.Position_X := 1500;
  SettingsChart.Position_Y := 5000;
  SettingsChart.StartRow := 0;
  SettingsChart.EndRow := 3;
  SettingsChart.PositionSheet := 0; // first tab
  SettingsChart.StartColumn := 'A';
  SettingsChart.EndColumn := 'B';
  SettingsChart.ChartName := 'TestChart';
  SettingsChart.typeChart := ctDefault;

  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctVertical;
  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctPie;
  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctLine;
  OpenOffice_calc1.addChart(SettingsChart);

  OpenOffice_writer1.saveFile('c:/temp/sheet.xlsx');
  OpenOffice_calc1.Free;
```


Writer(DOCX) Example
```pascal
var
  lTable: TTableWriter;
  OpenOffice_writer1:  TOpenOffice_writer;
begin
  OpenOffice_writer1 :=  TOpenOffice_writer.Create(Self);  
  OpenOffice_writer1.setBold(true).setFontHeight(16)
    .setValue(
       'Título: Apresentando o componente Libre Office writer via Delphi <3'+ #13#13)
    .gotoEndOfSentence
    .setBold(false).setFontHeight(12)
    .setValue(
       'Neste exemplo estou mostrando a criação de documentos via código, de um jeito simples, rápido e fácil.'+ #13)
    .gotoEndOfSentence
    .setBold(false).setFontHeight(12)
    .setValue(
        'Espero que seja útil e que estejam gostando, lembrando o componente é open source e totalmente free!'+ #13#13)
    .gotoEndOfSentence
    .setBold(false).setFontHeight(12).setBold(true)
      .setValue('Obrigado a todos pela presença!!!' + #13)
    .gotoEndOfSentence;

  lTable := TTableWriter.New;
  try
    lTable.TableFields.Add('Participantes');
    lTable.TableFields.Add('Idade');
    lTable.TableFields.Add('Local');

    lTable.Lines.Line.Add('Dinos');
    lTable.Lines.Line.Add('31');
    lTable.Lines.Line.Add('Brasil');

    lTable.Lines := TTableWriter.AddNewLines; //Adiciona uma nova linha
    lTable.Lines.Line.Add('Jao');
    lTable.Lines.Line.Add('22');
    lTable.Lines.Line.Add('USA');
    lTable.TableLines.Add(lTable.Lines);

    OpenOffice_writer1
      .CreateTable(lTable)
      .gotoEndOfSentence
      .setValue(#13#13);

    OpenOffice_writer1.saveFile('c:/temp/sdoc.docx');
  finally
    lTable.Free;
    OpenOffice_writer1.Free; 
  end;
end;
```

This software is MIT open source!

Wiki: https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/wiki

## See more on Instagram https://www.instagram.com/dinosdev/

### O componente é totalmente free, se ele foi muito útil para você, que tal me pagar um café para incentivar o projeto?

PIX:

<img src="https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/00dcc168-df75-4228-b80d-7262c7b4c478" width="300" height="300">


## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Daniel09Fernandes/DinosOffice&type=Date)](https://star-history.com/#Daniel09Fernandes/DinosOffice&Date)
