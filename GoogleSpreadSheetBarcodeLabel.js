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

                var sku = entry.gsx$sku.$t;
                var theme = entry.gsx$theme.$t;
		var imageurl = entry.gsx$imageurl.$t;

                var record = labelSet.addRecord();
                record.setText("SKU", sku);
                record.setText("THEME", theme);
		    
		
		//record.setText("BARCODE", imageurl);
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

            script.setAttribute('src', 'https://spreadsheets.google.com/feeds/list/1NuAevHWdpsWChNk20iWUYzTw9iJY1vRmnws7LSjGLv0/1/public/values?alt=json-in-script&callback=window._loadSpreadSheetDataCallback');
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
	<Id>FileFolder</Id>\
	<PaperName>30327 File Folder - offset</PaperName>\
	<DrawCommands>\
		<RoundRectangle X="0" Y="0" Width="806" Height="4950" Rx="180" Ry="180" />\
	</DrawCommands>\
	<ObjectInfo>\
		<ImageObject>\
			<Name>BARCODE</Name>\
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			<LinkedObjectName></LinkedObjectName>\
			<Rotation>Rotation0</Rotation>\
			<IsMirrored>False</IsMirrored>\
			<IsVariable>False</IsVariable>\
			<ImageLocation/>\
			<ScaleMode>Uniform</ScaleMode>\
			<BorderWidth>0</BorderWidth>\
			<BorderColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<HorizontalAlignment>Left</HorizontalAlignment>\
			<VerticalAlignment>Center</VerticalAlignment>\
		</ImageObject>\
		<Bounds X="316.799987792969" Y="57.6000137329102" Width="1111.20001220703" Height="691.200012207031" />\
	</ObjectInfo>\
	<ObjectInfo>\
		<TextObject>\
			<Name>SIZE</Name>\
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			<LinkedObjectName></LinkedObjectName>\
			<Rotation>Rotation0</Rotation>\
			<IsMirrored>False</IsMirrored>\
			<IsVariable>False</IsVariable>\
			<HorizontalAlignment>Left</HorizontalAlignment>\
			<VerticalAlignment>Top</VerticalAlignment>\
			<TextFitMode>None</TextFitMode>\
			<UseFullFontHeight>True</UseFullFontHeight>\
			<Verticalized>False</Verticalized>\
			<StyledText>\
				<Element>\
					<String>Beveled Edge: 100mm</String>\
					<Attributes>\
						<Font Family="Eurostile" Size="9" Bold="True" Italic="False" Underline="False" Strikeout="False" />\
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
					</Attributes>\
				</Element>\
			</StyledText>\
		</TextObject>\
		<Bounds X="1530" Y="57.6000137329102" Width="2865" Height="196.200012207031" />\
	</ObjectInfo>\
	<ObjectInfo>\
		<TextObject>\
			<Name>THEME</Name>\
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			<LinkedObjectName></LinkedObjectName>\
			<Rotation>Rotation0</Rotation>\
			<IsMirrored>False</IsMirrored>\
			<IsVariable>False</IsVariable>\
			<HorizontalAlignment>Left</HorizontalAlignment>\
			<VerticalAlignment>Top</VerticalAlignment>\
			<TextFitMode>ShrinkToFit</TextFitMode>\
			<UseFullFontHeight>True</UseFullFontHeight>\
			<Verticalized>False</Verticalized>\
			<StyledText>\
				<Element>\
					<String>Asian Garden</String>\
					<Attributes>\
						<Font Family="Eurostile" Size="9" Bold="True" Italic="False" Underline="False" Strikeout="False" />\
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
					</Attributes>\
				</Element>\
			</StyledText>\
		</TextObject>\
		<Bounds X="1530" Y="237.60001373291" Width="2880" Height="301.200012207031" />\
	</ObjectInfo>\
	<ObjectInfo>\
		<TextObject>\
			<Name>SKU</Name>\
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			<LinkedObjectName></LinkedObjectName>\
			<Rotation>Rotation0</Rotation>\
			<IsMirrored>False</IsMirrored>\
			<IsVariable>False</IsVariable>\
			<HorizontalAlignment>Left</HorizontalAlignment>\
			<VerticalAlignment>Top</VerticalAlignment>\
			<TextFitMode>ShrinkToFit</TextFitMode>\
			<UseFullFontHeight>True</UseFullFontHeight>\
			<Verticalized>False</Verticalized>\
			<StyledText>\
				<Element>\
					<String>(BAG1001)</String>\
					<Attributes>\
						<Font Family="Eurostile" Size="8" Bold="True" Italic="False" Underline="False" Strikeout="False" />\
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
					</Attributes>\
				</Element>\
			</StyledText>\
		</TextObject>\
		<Bounds X="1515" Y="432.60001373291" Width="1065" Height="211.200012207031" />\
	</ObjectInfo>\
	<ObjectInfo>\
		<TextObject>\
			<Name>URL</Name>\
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
			<LinkedObjectName></LinkedObjectName>\
			<Rotation>Rotation0</Rotation>\
			<IsMirrored>False</IsMirrored>\
			<IsVariable>False</IsVariable>\
			<HorizontalAlignment>Left</HorizontalAlignment>\
			<VerticalAlignment>Top</VerticalAlignment>\
			<TextFitMode>ShrinkToFit</TextFitMode>\
			<UseFullFontHeight>True</UseFullFontHeight>\
			<Verticalized>False</Verticalized>\
			<StyledText>\
				<Element>\
					<String>www.SecretWeaponMiniatures.com</String>\
					<Attributes>\
						<Font Family="Eurostile" Size="6" Bold="True" Italic="False" Underline="False" Strikeout="False" />\
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
					</Attributes>\
				</Element>\
			</StyledText>\
		</TextObject>\
		<Bounds X="1532.40008544922" Y="597.60001373291" Width="2880" Height="151.200012207031" />\
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

		//getImage();
		var imgurl = "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";
		label.setObjectText("BARCODE", imgurl);
                //label.print(printersSelect.value, '', labelSet);
                var records = labelSet.getRecords();
                for (var i = 0; i < records.length; ++i)
                {
                    label.setObjectText("SKU", records[i]["SKU"]);
                    label.setObjectText("THEME", records[i]["THEME"]);
//                    var pngData = label.render();

//                    var labelImage = document.getElementById('img' + (i + 1));
//                    labelImage.src = "data:image/png;base64," + pngData;
               	label.print(printersSelect.value);
		}

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
