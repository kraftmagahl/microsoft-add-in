
'use strict';


(function () {


    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                // Do something that is only available via the new APIs
                $('#fileUpload').change(ExcelToJSON);
                $('#uploadExcel').click(tables);
                $('#cat').click(funcCat);
                $('#ev').click(tables2);
                $('#graph').click(funcEv);

                $('#supportedVersion').html('This code is using Word 2016 or later.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or later.');
            }
        });
    });


    //duplica la parte selezionata del documento tante volte quante sono le entità studiate durante il pt.
    //Dati trovati nell'excel passato nell'input come file.

    function ExcelToJSON() {
        Word.run(function (context) {
            //recupero file excel dall'input
            var file = document.getElementById('fileUpload');
            if ('files' in file) {
                if (file.files.length == 0) {
                    context.document.body.insertText('no file selected', 'End');
                } else {
                    for (var i = 0; i < file.files.length; i++) {
                        var reader = new FileReader();
                        reader.onload = function (event) {
                            var data = reader.result;
                            var workbook = XLSX.read(data, {
                                type: "binary"
                            });
                            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[0]]);
                            var jsonObject = JSON.parse(JSON.stringify(rowObject));

                            //search where to change text and load it
                            let change1 = context.document.body.search('ab12');
                            let change2 = context.document.body.search('abc123');
                            context.load(change1);
                            context.load(change2);

                            //build the string with which switch the text 
                            var string1 = "";
                            var string2 = "";
                            for (var i = 0; i < jsonObject.length; i++) {
                                string1 += ("\nEntità: " + jsonObject[i]['N'] + ": " + jsonObject[i]['Target/URL']);
                                string2 += ((i + 1) + ".  " + jsonObject[i]['Target/URL'] + "\n");
                            }

                            //create as many tables as I need
                            const table = context.document.getSelection();
                            var stringa = table.getHtml();
                            table.delete();

                            return context.sync().then(() => {
                                for (var i = 0; i < jsonObject.length; i++) {
                                    table.insertHtml(stringa.value + "<br />", 'End');
                                }
                                //switch text 
                                change1.items.forEach(function (obj) {
                                    obj.insertText(string1, 'Replace');
                                });
                                change2.items.forEach(function (obj) {
                                    obj.insertText(string2, 'Replace');
                                });
                            }).then(context.sync);

                        };
                        reader.readAsBinaryString(file.files[i], "UTF-8");
                    }
                }
            } else {
                context.document.body.insertText('no file found', 'End');
            }


            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    //compila i nomi e url delle entità studiate 
    function tables() {
        Word.run(function (context) {
            //recupero file excel dall'input
            var file = document.getElementById('fileUpload');
            if ('files' in file) {
                if (file.files.length == 0) {
                    context.document.body.insertText('no file selected', 'End');
                } else {
                    for (var i = 0; i < file.files.length; i++) {
                        var reader = new FileReader();
                        reader.onload = function (event) {
                            var data = reader.result;
                            var workbook = XLSX.read(data, {
                                type: "binary"
                            });
                            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[0]]);
                            var jsonObject = JSON.parse(JSON.stringify(rowObject));
                            //find words to be changed and load them
                            let change3 = context.document.body.search('EX*', {
                                matchWildCards: true
                            });
                            let change4 = context.document.body.search('x.x.x.x.');
                            context.load(change3);
                            context.load(change4);

                            return context.sync().then(() => {
                                //change text of the words found
                                var j = 1;
                                for (var i = 0; i < jsonObject.length; i++) {
                                    change3.items[j].insertText(jsonObject[i]['N'], 'Replace');
                                    j++;
                                    change3.items[j].insertText(jsonObject[i]['N'], 'Replace');
                                    j++;
                                    change4.items[i + 1].insertText(jsonObject[i]['Target/URL'], 'Replace');
                                }
                            }).then(context.sync);
                        };
                        reader.readAsBinaryString(file.files[i], "UTF-8");
                    }
                }
            } else {
                context.document.body.insertText('no file found', 'End');
            }

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    //duplica la parte selezionata tante volte quanti sono i tipi di vulnerabilità riscontrati per l'entità 
    //definita nell'input text. Aggiunge anche le tabelle delle vulnerabilità con il numero di risorse vulnerabili definito. 
    //Dati trovati nell'excel passato nell'input come file.
    function funcCat() {
        Word.run(function (context) {
            //recupero excel dall'input
            var categoria = document.getElementById('inputcat');
            var file = document.getElementById('fileUpload');
            if ('files' in file) {
                if (file.files.length == 0) {
                    context.document.body.insertText('no file selected', 'End');
                } else {
                    for (var i = 0; i < file.files.length; i++) {
                        var reader = new FileReader();
                        reader.onload = function (event) {
                            var data = reader.result;
                            var workbook = XLSX.read(data, {
                                type: "binary"
                            });
                            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[1]]);
                            var jsonObject = JSON.parse(JSON.stringify(rowObject));

                            //create as many tables as I need
                            const table = context.document.getSelection();
                            var cont = 0;
                            var prec = "";
                            for (var i = 0; i < jsonObject.length; i++) {
                                if (jsonObject[i]['Target'] == categoria.value && jsonObject[i]['Cat. OWASP'] != prec) {
                                    cont++;
                                    prec = jsonObject[i]['Cat. OWASP'];
                                }
                            }
                            var stringa = table.getHtml();
                            table.delete();

                            return context.sync().then(() => {
                                if (cont != 0) {
                                    var exist = 0;
                                    for (var i = 0; i < cont; i++) {
                                        table.insertHtml(stringa.value, 'End');
                                        while (jsonObject[exist]['Target'] != categoria.value) {
                                            exist++;
                                        }
                                        //find number of vulnerable resources for the vullnerability studied
                                        var vuln = jsonObject[exist]['Risorse \r\nVuln.'];
                                        if (exist + 1 < jsonObject.length) {
                                            while (jsonObject[exist]['Cat. OWASP'] == jsonObject[exist + 1]['Cat. OWASP']) {
                                                vuln += jsonObject[exist + 1]['Risorse \r\nVuln.'];
                                                exist++;
                                            }
                                        }
                                        //insert the corrisponding table of vulnerability 
                                        if (jsonObject[exist]['Gravità'] == 'BASSA') {
                                            table.insertHtml('<table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="643" style="border-collapse:collapse;mso-table-layout-alt:fixed;border:none; mso-yfti-tbllook:1184;mso-padding-alt:5.65pt 5.4pt 5.65pt 5.4pt;mso-border-insideh:none;mso-border-insidev:none"><tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:21.4pt"><td width="473" colspan="2" valign="top" style="width:354.4pt;border:none;border-bottom:solid #C2D72C 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-TITOLO4-VULNCRITICA" style="margin-left:-5.45pt;mso-add-space:auto"><span style="color:#C2D72C">Nome Evidenza (CWE X)<o:p></o:p></span></p></td><td width="170" valign="top" style="width:127.7pt;border:none;border-bottom:solid #C2D72C 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-NORMALE" align="right" style="margin-right:-5.15pt;mso-add-space:auto;text-align:right"><span style="font-size:12.0pt;mso-bidi-font-size:14.0pt;line-height:115%;mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#C2D72C">Gravità: <b style="mso-bidi-font-weight:normal">BASSA</b></span><span style="font-size:12.0pt;mso-bidi-font-size:10.0pt;line-height:115%;mso-bidi-font-family:Arial;color:#C2D72C"><o:p></o:p></span></p></td></tr><tr style="mso-yfti-irow:1;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;border:none;mso-border-top-alt:solid #C2D72C 1.5pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">DESCRIZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;border:none;mso-border-top-alt:solid #C2D72C 1.5pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpLast"><span style="color:black;mso-color-alt:windowtext">Descrizione</span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p></td></tr><tr style="mso-yfti-irow:2;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">SOLUZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpLast"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D">Soluzione </span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:3;height:19.75pt"> <td width="208" valign="top" style="width:155.95pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">Risorse Vulnerabili (' + vuln + ')<o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpMiddle" style="margin-left:36.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol"><span style="mso-list:Ignore">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">URL:</span><o:p></o:p></p><p class="Y-NORMALECxSpLast" style="margin-left:72.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level2 lfo1"><!--[if !supportLists]--><span style="font-family:&quot;Courier New&quot;;mso-fareast-font-family:&quot;Courier New&quot;"><span style="mso-list:Ignore">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">Parametro:</span><o:p></o:p></p> </td></tr><tr style="mso-yfti-irow:4;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:white;mso-background-themecolor background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">PROFILO<o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpLast"><span style="color:red">Abilitazioni di rete standard / concesse</span><o:p></o:p></p></td></tr><tr style="mso-yfti-irow:5;height:19.75pt"><td width="643" colspan="3" valign="top" style="width:482.1pt;background:#F2F2F2;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALE" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">PROOF OF CONCEPT<o:p></o:p></span></b></p></td></tr><tr style="mso-yfti-irow:6;mso-yfti-lastrow:yes;height:19.75pt"><td width="643" colspan="3" valign="top" style="width:482.1pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALE"><span style="color:black;mso-color-alt:windowtext">Proof of concept</span><o:p></o:p></p></td></tr></tbody></table><p class="Y-NORMALE"><o:p>&nbsp;</o:p></p>', 'End');

                                        } if (jsonObject[exist]['Gravità'] == 'MEDIA') {
                                            table.insertHtml('<table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="643" style="border-collapse:collapse;mso-table-layout-alt:fixed;border:none;mso-yfti-tbllook:1184;mso-padding-alt:5.65pt 5.4pt 5.65pt 5.4pt;mso-border-insideh:none;mso-border-insidev:none"<tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:21.4pt"><td width="473" colspan="2" valign="top" style="width:354.4pt;border:none;border-bottom:solid #FFC000 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-TITOLO4-VULNCRITICA" style="margin-left:-5.45pt;mso-add-space:auto"><span style="color:#FFC000">Nome Evidenza (CWE X)<o:p></o:p></span></p></td><td width="170" valign="top" style="width:127.7pt;border:none;border-bottom:solid #FFC000 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-NORMALE" align="right" style="margin-right:-5.15pt;mso-add-space:auto;text-align:right"><span style="font-size:12.0pt;mso-bidi-font-size:14.0pt;line-height:115%;mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#FFC000">Gravità: <b style="mso-bidi-font-weight:normal">MEDIA</b></span><span style="font-size:12.0pt;mso-bidi-font-size:10.0pt;line-height:115%;mso-bidi-font-family:Arial;color:#FFC000"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:1;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;border:none;mso-border-top-alt:solid #FFC000 1.5pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">DESCRIZIONE</span></b><b style="mso-bidi-font-weight: normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p> </td> <td width="435" colspan="2" valign="top" style="width:326.15pt;border:none; mso-border-top-alt:solid #FFC000 1.5pt;background:#F2F2F2;mso-background-themecolor: background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt; height:19.75pt"> <p class="Y-NORMALECxSpLast"><span style="color:black;mso-color-alt:windowtext">Descrizione</span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:2;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space: auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">SOLUZIONE</span></b><b style="mso-bidi-font-weight: normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p> </td> <td width="435" colspan="2" valign="top" style="width:326.15pt;background:white; mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt; height:19.75pt"> <p class="Y-NORMALECxSpLast"><span style="mso-fareast-font-family:Calibri; mso-bidi-font-family:Arial;color:#101F2D">Soluzione </span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:3;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">Risorse Vulnerabili (' + vuln + ')<o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:#F2F2F2; mso-background-themecolor:background1;mso-background-themeshade:242; padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpMiddle" style="margin-left:18.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol"><span style="mso-list:Ignore">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">URL:</span><o:p></o:p></p><p class="Y-NORMALECxSpLast" style="margin-left:72.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level2 lfo1"><!--[if !supportLists]--><span style="font-family:&quot;Courier New&quot;;mso-fareast-font-family:&quot;Courier New&quot;"><span style="mso-list:Ignore">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">Parametro:</span><o:p></o:p></p></td></tr><tr style="mso-yfti-irow:4;height:19.75pt"><td width="643" colspan="3" valign="top" style="width:482.1pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALE" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">PROOF OF CONCEPT<o:p></o:p></span></b></p> </td></tr><tr style="mso-yfti-irow:5;mso-yfti-lastrow:yes;height:19.75pt"><td width="643" colspan="3" valign="top" style="width:482.1pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALE"><span style="color:black;mso-color-alt:windowtext">Proof of concept</span><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family: Arial;color:#101F2D"><o:p></o:p></span></p></td></tr><!--[if !supportMisalignedColumns]--><tr height="0"><td width="208" style="border:none"></td><td width="265" style="border:none"></td><td width="170" style="border:none"></td></tr> <!--[endif]--></tbody></table><p class="Y-NORMALE"><o:p>&nbsp;</o:p></p><span style="font-size:10.0pt;line-height:107%;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:minor-bidi;mso-ansi-language:IT;mso-fareast-language:IT;mso-bidi-language:AR-SA"><br clear="all" style="mso-special-character:line-break;page-break-before:always"></span><p class="MsoNormal" align="left" style="margin-bottom:8.0pt;text-align:left;line-height:107%;mso-hyphenate:auto"><b><span style="font-size:12.0pt;line-height:107%;color:#0055B8;mso-fareast-language:EN-US"><o:p>&nbsp;</o:p></span></b></p>', 'End');

                                        } if (jsonObject[exist]['Gravità'] == 'ALTA') {
                                            table.insertHtml('<table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="643" style="border-collapse:collapse;mso-table-layout-alt:fixed;border:none;mso-yfti-tbllook:1184;mso-padding-alt:5.65pt 5.4pt 5.65pt 5.4pt;mso-border-insideh:none;mso-border-insidev:none"><tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:21.4pt"><td width="473" colspan="2" valign="top" style="width:354.4pt;border:none;border-bottom:solid red 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-TITOLO4-VULNCRITICA" style="margin-left:-5.45pt;mso-add-space:auto"><span style="color:red">Nome Evidenza (CWE X)<o:p></o:p></span></p></td><td width="170" valign="top" style="width:127.7pt;border:none;border-bottom:solid red 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-NORMALE" align="right" style="margin-right:-5.15pt;mso-add-space:auto;text-align:right"><span style="font-size:12.0pt;mso-bidi-font-size:14.0pt;line-height:115%;mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:red">Gravità: <b style="mso-bidi-font-weight:normal">ALTA</b></span><span style="font-size:12.0pt;mso-bidi-font-size:10.0pt;line-height:115%;mso-bidi-font-family:Arial;color:red"><o:p></o:p></span></p></td></tr><tr style="mso-yfti-irow:1;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;border:none;mso-border-top-alt:solid red 1.5pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">DESCRIZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;border:none;mso-border-top-alt:solid #7030A0 1.5pt;background:#F2F2F2;mso-background-themecolor: background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALECxSpLast"><span style="color:black;mso-color-alt:windowtext">Descrizione</span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p></td></tr><tr style="mso-yfti-irow:2;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">SOLUZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpLast"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D">Soluzione </span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:3;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:#F2F2F2;mso-background-themecolor: background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">Risorse Vulnerabili (' + vuln + ')<o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpMiddle" style="margin-left:18.0pt;mso-add-space:auto; text-indent:-18.0pt;mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family: Symbol"><span style="mso-list:Ignore">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">URL:</span><o:p></o:p></p><p class="Y-NORMALECxSpLast" style="margin-left:72.0pt;mso-add-space:auto; text-indent:-18.0pt;mso-list:l0 level2 lfo1"><!--[if !supportLists]--><span style="font-family:&quot;Courier New&quot;;mso-fareast-font-family:&quot;Courier New&quot;"><span style="mso-list:Ignore">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp; </span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">Parametro:</span><o:p></o:p></p></td></tr><tr style="mso-yfti-irow:4;height:19.75pt"><td width="643" colspan="3" valign="top" style="width:482.1pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALE" align="left" style="margin-left:-5.45pt;mso-add-space:auto; text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">PROOF OF CONCEPT<o:p></o:p></span></b></p> </td></tr><tr style="mso-yfti-irow:5;mso-yfti-lastrow:yes;height:19.75pt"> <td width="643" colspan="3" valign="top" style="width:482.1pt;background:white; mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALE"><span style="color:black;mso-color-alt:windowtext">Proof of concept</span><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D"><o:p></o:p></span></p></td></tr><!--[if !supportMisalignedColumns]--><tr height="0"><td width="208" style="border:none"></td><td width="265" style="border:none"></td><td width="170" style="border:none"></td></tr> <!--[endif]--></tbody></table><p class="Y-NORMALE"><o:p>&nbsp;</o:p></p><span style="font-size:10.0pt;line-height:107%;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:minor-bidi;mso-ansi-language:IT;mso-fareast-language:IT;mso-bidi-language:AR-SA"><br clear="all" style="mso-special-character:line-break;page-break-before:always"></span><p class="MsoNormal" align="left" style="margin-bottom:8.0pt;text-align:left;line-height:107%;mso-hyphenate:auto"><b><span style="font-size:12.0pt;line-height:107%;color:#0055B8;mso-fareast-language:EN-US"><o:p>&nbsp;</o:p></span></b></p>', 'End');

                                        } if (jsonObject[exist]['Gravità'] == 'GRAVE') {
                                            table.insertHtml('<table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="643" style="border-collapse:collapse;mso-table-layout-alt:fixed;border:none; mso-yfti-tbllook:1184;mso-padding-alt:5.65pt 5.4pt 5.65pt 5.4pt;mso-border-insideh: none;mso-border-insidev:none"> <tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:21.4pt"> <td width="473" colspan="2" valign="top" style="width:354.4pt;border:none; border-bottom:solid #7030A0 1.5pt;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-TITOLO4-VULNCRITICA" style="margin-left:-5.45pt;mso-add-space:auto">Nome Evidenza (CWE X)<o:p></o:p></p></td> <td width="170" valign="top" style="width:127.7pt;border:none;border-bottom:solid #7030A0 1.5pt; padding:5.65pt 5.4pt 5.65pt 5.4pt;height:21.4pt"><p class="Y-NORMALE" align="right" style="margin-right:-5.15pt;mso-add-space: auto;text-align:right"><span style="font-size:12.0pt;mso-bidi-font-size:14.0pt;line-height:115%;mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial; color:#7030A0">Gravità: <b style="mso-bidi-font-weight:normal">CRITICA</b></span><span style="font-size:12.0pt;mso-bidi-font-size:10.0pt;line-height:115%; mso-bidi-font-family:Arial;color:#7030A0"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:1;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;border:none;mso-border-top-alt:solid #7030A0 1.5pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space: auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">DESCRIZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;border:none;mso-border-top-alt:solid windowtext .5pt;background:#F2F2F2;mso-background-themecolor: background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt; height:19.75pt"><p class="Y-NORMALECxSpLast"><span style="color:black;mso-color-alt:windowtext">Descrizione</span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p></td></tr><tr style="mso-yfti-irow:2;height:19.75pt"><td width="208" valign="top" style="width:155.95pt;background:white;mso-background-themecolor: background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space: auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D;text-transform:uppercase">SOLUZIONE</span></b><b style="mso-bidi-font-weight:normal"><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></b></p></td><td width="435" colspan="2" valign="top" style="width:326.15pt;background:white;mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALECxSpLast"><span style="mso-fareast-font-family:Calibri; mso-bidi-font-family:Arial;color:#101F2D">Soluzione </span><span style="mso-bidi-font-family:Arial"><o:p></o:p></span></p> </td></tr><tr style="mso-yfti-irow:3;height:19.75pt"> <td width="208" valign="top" style="width:155.95pt;background:#F2F2F2;mso-background-themecolor:background1;mso-background-themeshade:242;padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"><p class="Y-NORMALECxSpFirst" align="left" style="margin-left:-5.45pt;mso-add-space:auto;text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">Risorse Vulnerabili (' + vuln + ')<o:p></o:p></span></b></p> </td> <td width="435" colspan="2" valign="top" style="width:326.15pt;background:#F2F2F2; mso-background-themecolor:background1;mso-background-themeshade:242; padding:5.65pt 5.4pt 5.65pt 5.4pt;height:19.75pt"> <p class="Y-NORMALECxSpMiddle" style="margin-left:18.0pt;mso-add-space:auto; text-indent:-18.0pt;mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol"><span style="mso-list:Ignore">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">URL:</span><o:p></o:p></p><p class="Y-NORMALECxSpLast" style="margin-left:72.0pt;mso-add-space:auto; text-indent:-18.0pt;mso-list:l0 level2 lfo1"><!--[if !supportLists]--><span style="font-family:&quot;Courier New&quot;;mso-fareast-font-family:&quot;Courier New&quot;"><span style="mso-list:Ignore">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;</span></span></span><!--[endif]--><span style="color:black;mso-color-alt:windowtext">Parametro:</span><o:p></o:p></p> </td></tr><tr style="mso-yfti-irow:4;height:19.75pt"> <td width="643" colspan="3" valign="top" style="width:482.1pt;background:white; mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt; height:19.75pt"> <p class="Y-NORMALE" align="left" style="margin-left:-5.45pt;mso-add-space:auto; text-align:left"><b style="mso-bidi-font-weight:normal"><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D; text-transform:uppercase">PROOF OF CONCEPT<o:p></o:p></span></b></p> </td></tr><tr style="mso-yfti-irow:5;mso-yfti-lastrow:yes;height:19.75pt"> <td width="643" colspan="3" valign="top" style="width:482.1pt;background:white; mso-background-themecolor:background1;padding:5.65pt 5.4pt 5.65pt 5.4pt; height:19.75pt"> <p class="Y-NORMALE"><span style="color:black;mso-color-alt:windowtext">Proof of concept</span><span style="mso-fareast-font-family:Calibri;mso-bidi-font-family:Arial;color:#101F2D"><o:p></o:p></span></p></td></tr><!--[if !supportMisalignedColumns]--><tr height="0"><td width="208" style="border:none"></td><td width="265" style="border:none"></td><td width="170" style="border:none"></td></tr><!--[endif]--></tbody></table><p class="Y-NORMALE"><o:p>&nbsp;</o:p></p>', 'End');

                                        }
                                        exist++;
                                    }

                                } else {
                                    table.insertHtml("<br />" + "<br />" + "<p>Nessuna vulnerabilità rilevata per il target indicato. Target non raggiungibile.</p>" + "<br />" + "<br />", 'End');
                                }
                            }).then(context.sync);

                        };
                        reader.readAsBinaryString(file.files[i], "UTF-8");
                    }
                }
            } else {
                context.document.body.insertText('no file found', 'End');
            }


            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });

    }

    //Inserisce le categorie di vulnerabilità.
    //compila le tabelle delle vulnerabilità dell'entità definita nell'input text con i titoli corrispondenti.
    function tables2() {
        Word.run(function (context) {
            //recupero file excel dall'input
            var categoria = document.getElementById('inputcat');
            var file = document.getElementById('fileUpload');
            if ('files' in file) {
                if (file.files.length == 0) {
                    context.document.body.insertText('no file selected', 'End');
                } else {
                    for (var i = 0; i < file.files.length; i++) {
                        var reader = new FileReader();
                        reader.onload = function (event) {
                            var data = reader.result;
                            var workbook = XLSX.read(data, {
                                type: "binary"
                            });
                            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[1]]);
                            var jsonObject = JSON.parse(JSON.stringify(rowObject));

                            //var prova = descr(JSON.stringify(jsonObject[0]['Titolo vulnerabilità (CWE-X)']));//////////////////////////////////////////////////
                            var prova = descr("wowowowowowowow");
                            //search words to be changed and load them
                            let change1 = context.document.body.search('OWASP A1 – Injection');
                            let change2 = context.document.body.search('Nome Evidenza (CWE X)');
                            let change3 = context.document.body.search('.A1.1');
                            context.load(change1);
                            context.load(change2);
                            context.load(change3);
                            //change the words found with data found in the excel
                            return context.sync().then(() => {
                                var prec = "";
                                var j = 0;
                                var n = 1;
                                for (var i = 0; i < jsonObject.length; i++) {
                                    if (jsonObject[i]['Target'] == categoria.value && jsonObject[i]['Cat. OWASP'] != prec) {
                                        change1.items[j].insertText(jsonObject[i]['Cat. OWASP'], 'Replace');
                                        var text = jsonObject[i]['Titolo vulnerabilità (CWE-X)'];
                                        change2.items[j].insertText("" + text.toUpperCase(), 'Replace');
                                        change3.items[j].insertText("." + jsonObject[i]['Cat. OWASP'].substr(0, 2) + "." + n, 'Replace');
                                        j++;
                                        prec = jsonObject[i]['Cat. OWASP'];
                                    }
                                }
                                context.document.body.insertText(JSON.stringify(prova) + "", 'End');/////////////////////////////////////////////////////////////////
                            }).then(context.sync);

                        };
                        reader.readAsBinaryString(file.files[i], "UTF-8");
                    }
                }
            } else {
                context.document.body.insertText('no file found', 'End');
            }


            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });


    }

    //inserisce dati nella tabella della distribuzione complessiva delle vulnerabilità rilevate.
    function funcEv() {
        Word.run(function (context) {
            //recupero excel dall'input
            var file = document.getElementById('fileUpload');
            if ('files' in file) {
                if (file.files.length == 0) {
                    context.document.body.insertText('no file selected', 'End');
                } else {
                    for (var i = 0; i < file.files.length; i++) {
                        var reader = new FileReader();
                        reader.onload = function (event) {
                            var data = reader.result;
                            var workbook = XLSX.read(data, {
                                type: "binary"
                            });
                            var sheet = workbook.Sheets[workbook.SheetNames[3]];
                            let rowObject = XLSX.utils.sheet_to_row_object_array(sheet);
                            var jsonObject = JSON.parse(JSON.stringify(rowObject));
                            //conto le vulnerabilità univoche 
                            var basse = 0;
                            var medie = 0;
                            var alte = 0;
                            var critiche = 0;
                            for (var i = 0; i < 11; i++) {
                                if (jsonObject[i]['Critiche'] != 0) {
                                    critiche++;
                                }
                                if (jsonObject[i]['Medie'] != 0) {
                                    medie++;
                                }
                                if (jsonObject[i]['Basse'] != 0) {
                                    basse++;
                                }
                                if (jsonObject[i]['Alte'] != 0) {
                                    alte++;
                                }
                            }
                            //search the word to be changed and load
                            var cose = context.document.body.search('X?');
                            context.load(cose);
                            //change the words found in the document with excel data
                            return context.sync().then(() => {
                                cose.items[1].insertText(jsonObject[11]['Critiche'] + "", 'Replace');
                                cose.items[3].insertText(jsonObject[11]['Alte'] + "", 'Replace');
                                cose.items[5].insertText(jsonObject[11]['Medie'] + "", 'Replace');
                                cose.items[7].insertText(jsonObject[11]['Basse'] + "", 'Replace');
                                cose.items[0].insertText(critiche + "", 'Replace');
                                cose.items[2].insertText(alte + "", 'Replace');
                                cose.items[4].insertText(medie + "", 'Replace');
                                cose.items[6].insertText(basse + "", 'Replace');

                            }).then(context.sync);

                        };
                        reader.readAsBinaryString(file.files[i], "UTF-8");
                    }
                }
            } else {
                context.document.body.insertText('no file found', 'End');
            }


            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });


    }

    function descr(titolo) {
        var res = "0";
        var http = new XMLHttpRequest();
        http.onreadystatechange = function () {
            if (http.readyState == 4 && http.status == 200) {
                res = http.responseText;
            } else {
                res = http.statusText + " " + http.readyState;
            }
        };
        http.withCredentials = true;
        //http.setRequestHeader('Access-Control-Allow-Origin','*');
        //http.setRequestHeader("Content-type", "text/plain"); 
        http.open("GET", "https://f8e9fa35212d.ngrok.io/api/json/" + titolo, true);
        http.send(null);
        return "";
    }
})();