//----------------------------------------------------------------------------
//
//  $Id: GoogleSpreadSheetBarcodeLabel.js 38773 2015-09-17 11:45:41Z nmikalko $ 
//
// Project -------------------------------------------------------------------
//
//  DYMO Label Framework
//
// Content -------------------------------------------------------------------
//
//  DYMO Label Framework JavaScript Library Samples: 
//    Print mulltiple labels using Google Spreadsheet as a data source
//
//----------------------------------------------------------------------------
//
//  Copyright (c), 2010, Sanford, L.P. All Rights Reserved.
//
//----------------------------------------------------------------------------



(function()
{
    var label;
    var labelSet;

    function onload()
    {
        var printButton = document.getElementById('printButton');
        var printersSelect = document.getElementById('printersSelect');

        function createLabelSet(json)
        {
            var labelSet = new dymo.label.framework.LabelSetBuilder();
         
            for (var i = 0; i < json.feed.entry.length; ++i)
            {
                var entry = json.feed.entry[i];

                var itemDescription = entry.gsx$itemdescription.$t;
                var itemCode = entry.gsx$itemcode.$t;

                var record = labelSet.addRecord();
                record.setText("Description", itemDescription);
                record.setText("ItemCode", itemCode);
            }

            return labelSet;
        }

        function loadSpreadSheetDataCallback(json)
        {
            labelSet = createLabelSet(json);
        };

        window._loadSpreadSheetDataCallback = loadSpreadSheetDataCallback;

        function loadSpreadSheetData()
        {
            removeOldJSONScriptNodes();

            var script = document.createElement('script');

            script.setAttribute('src', 'https://spreadsheets.google.com/feeds/cells/1NuAevHWdpsWChNk20iWUYzTw9iJY1vRmnws7LSjGLv0/1/public/full?alt=json');
            script.setAttribute('id', 'printScript');
            script.setAttribute('type', 'text/javascript');
            document.documentElement.firstChild.appendChild(script);
        };

        function removeOldJSONScriptNodes()
        {
            var jsonScript = document.getElementById('printScript');
            if (jsonScript)
                jsonScript.parentNode.removeChild(jsonScript);
        };

        function getBarcodeLabelXml()
        {

            var labelXml = '<?xml version="1.0" encoding="utf-8"?>\
<DieCutLabel Version="8.0" Units="twips">\
	                            <PaperOrientation>Landscape</PaperOrientation>\
	                            <Id>Address</Id>\
	                            <PaperName>30252 Address</PaperName>\
	                            <DrawCommands>\
		                            <RoundRectangle X="0" Y="0" Width="1581" Height="5040" Rx="270" Ry="270" />\
	                            </DrawCommands>\
	                            <ObjectInfo>\
		                            <AddressObject>\
			                            <Name>Address</Name>\
			                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			                            <BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			                            <LinkedObjectName></LinkedObjectName>\
			                            <Rotation>Rotation0</Rotation>\
			                            <IsMirrored>False</IsMirrored>\
			                            <IsVariable>True</IsVariable>\
			                            <HorizontalAlignment>Left</HorizontalAlignment>\
			                            <VerticalAlignment>Middle</VerticalAlignment>\
			                            <TextFitMode>ShrinkToFit</TextFitMode>\
			                            <UseFullFontHeight>True</UseFullFontHeight>\
			                            <Verticalized>False</Verticalized>\
			                            <StyledText>\
				                            <Element>\
					                            <String>DYMO\
                                        828 San Pablo Ave Ste 101\
                                        Albany, CA 94706-1678</String>\
                                                            <Attributes>\
                                                                <Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
                                                                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
                                                            </Attributes>\
                                                        </Element>\
                                                    </StyledText>\
                                                    <ShowBarcodeFor9DigitZipOnly>False</ShowBarcodeFor9DigitZipOnly>\
                                                    <BarcodePosition>AboveAddress</BarcodePosition>\
                                                    <LineFonts>\
                                                        <Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
                                                        <Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
                                                        <Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
                                                    </LineFonts>\
                                     </AddressObject>\
                                     <Bounds X="332" Y="150" Width="4455" Height="1260" />\
                                  </ObjectInfo>\
                             </DieCutLabel>';
            return labelXml;
        }

        function loadLabel()
        {
            // use jQuery API to load label
            //$.get("Barcode.label", function(labelXml)
            //{
            label = dymo.label.framework.openLabelXml(getBarcodeLabelXml());
            //}, "text");
        }

        // loads all supported printers into a combo box 
        function loadPrinters()
        {
            var printers = dymo.label.framework.getLabelWriterPrinters();
            if (printers.length == 0)
            {
                alert("No DYMO LabelWriter printers are installed. Install DYMO LabelWriter printers.");
                return;
            }

            for (var i = 0; i < printers.length; ++i)
            {
                var printer = printers[i];
                var printerName = printer.name;

                var option = document.createElement('option');
                option.value = printerName;
                option.appendChild(document.createTextNode(printerName));
                printersSelect.appendChild(option);
            }
        }

        // prints the label
        printButton.onclick = function()
        {
            try
            {
                if (!label)
                    throw "Label is not loaded";

                if (!labelSet)
                    throw "Label data is not loaded";

                label.print(printersSelect.value, '', labelSet);

//                var records = labelSet.getRecords();
//                for (var i = 0; i < records.length; ++i)
//                {
//                    label.setObjectText("Description", records[i]["Description"]);
//                    label.setObjectText("ItemCode", records[i]["ItemCode"]);
//                    var pngData = label.render();
//
//                    var labelImage = document.getElementById('img' + (i + 1));
//                    labelImage.src = "data:image/png;base64," + pngData;
//                }
            }
            catch (e)
            {
                alert(e.message || e);
            }
        };

        loadLabel();
        loadSpreadSheetData();
        loadPrinters();

    };

    function initTests()
	{
		if(dymo.label.framework.init)
		{
			//dymo.label.framework.trace = true;
			dymo.label.framework.init(onload);
		} else {
			onload();
		}
	}

	// register onload event
	if (window.addEventListener)
		window.addEventListener("load", initTests, false);
	else if (window.attachEvent)
		window.attachEvent("onload", initTests);
	else
		window.onload = initTests;

} ());
