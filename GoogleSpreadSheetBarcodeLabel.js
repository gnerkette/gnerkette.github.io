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

            script.setAttribute('src', 'http://spreadsheets.google.com/feeds/list/0Ak2RtsSi0A2bdFdka0Y1VUtZZHQ0VlRGQXg5QzROb2c/1/public/values?alt=json-in-script&callback=window._loadSpreadSheetDataCallback');
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
		                            <BarcodeObject>\
			                            <Name>ItemCode</Name>\
			                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			                            <BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			                            <LinkedObjectName></LinkedObjectName>\
			                            <Rotation>Rotation0</Rotation>\
			                            <IsMirrored>False</IsMirrored>\
			                            <IsVariable>True</IsVariable>\
			                            <Text>1234</Text>\
			                            <Type>Code128Auto</Type>\
			                            <Size>Small</Size>\
			                            <TextPosition>Bottom</TextPosition>\
			                            <TextFont Family="Arial" Size="6" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
			                            <CheckSumFont Family="Arial" Size="7.3125" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
			                            <TextEmbedding>None</TextEmbedding>\
			                            <ECLevel>0</ECLevel>\
			                            <HorizontalAlignment>Center</HorizontalAlignment>\
			                            <QuietZonesPadding Left="0" Top="0" Right="0" Bottom="0" />\
		                            </BarcodeObject>\
		                            <Bounds X="331" Y="680.31494140625" Width="4622" Height="765.708679199219" />\
	                            </ObjectInfo>\
	                            <ObjectInfo>\
		                            <TextObject>\
			                            <Name>Description</Name>\
			                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			                            <BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			                            <LinkedObjectName></LinkedObjectName>\
			                            <Rotation>Rotation0</Rotation>\
			                            <IsMirrored>False</IsMirrored>\
			                            <IsVariable>True</IsVariable>\
			                            <HorizontalAlignment>Center</HorizontalAlignment>\
			                            <VerticalAlignment>Top</VerticalAlignment>\
			                            <TextFitMode>ShrinkToFit</TextFitMode>\
			                            <UseFullFontHeight>True</UseFullFontHeight>\
			                            <Verticalized>False</Verticalized>\
			                            <StyledText>\
				                            <Element>\
					                            <String>ItemDescription</String>\
					                            <Attributes>\
						                            <Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />\
						                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
					                            </Attributes>\
				                            </Element>\
			                            </StyledText>\
		                            </TextObject>\
		                            <Bounds X="331" Y="163" Width="4622" Height="341.566925048828" />\
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
