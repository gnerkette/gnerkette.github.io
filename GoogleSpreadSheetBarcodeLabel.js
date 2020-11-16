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
    var upcimages;
	var current= 0;

    function onload()
    {
        var printButton = document.getElementById('printButton');
        var printersSelect = document.getElementById('printersSelect');

        function createLabelSet(json)
        {
            var labelSet = new dymo.label.framework.LabelSetBuilder();
         
            for (var i = current; i < json.feed.entry.length; ++i)
            {
                var entry = json.feed.entry[i];

                var sku = entry.gsx$sku.$t;
                var theme = entry.gsx$theme.$t;
		var upc = entry.gsx$upcimage.$t;

                var record = labelSet.addRecord();
                record.setText("SKU", sku);
                record.setText("THEME", theme);
		record.setText("BARCODE", upc);
		    

		//upcimages[i] = record.setText("BARCODE", upc);
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

            for (var i = current; i < printers.length; ++i)
            {
                var printer = printers[i];
		    alert(printers[i]);
                var printerName = printer.name;

                var option = document.createElement('option');
                option.value = printerName;
                option.appendChild(document.createTextNode(printerName));
                printersSelect.appendChild(option);
            }
        }

        // prints the label
       function printnow()
        {
	   var i = 0;
            try
            {
                if (!label)
                    throw "Label is not loaded";

                if (!labelSet)
                    throw "Label data is not loaded";

         //       label.print(printersSelect.value, '', labelSet);

                var records = labelSet.getRecords();
		    alert(records[0]["SKU"]);
                for (i=0; i < records.length; ++i)
                {
		  	label.setObjectText('THEME', records[i]["THEME"]);
			label.setObjectText('SKU', records[i]["SKU"]);	
			
			
			
		   var img = new Image();
                img.crossOrigin = 'anonymous';
                img.onload = function()
                {
                    try
                    {
                        var canvas = document.createElement('canvas');
                        canvas.width = img.width;                     
                        canvas.height = img.height;

                        var context = canvas.getContext('2d');
                        context.drawImage(img, 0, 0);

                        var dataUrl = canvas.toDataURL('image/png');
                        var pngBase64 = dataUrl.substr('data:image/png;base64,'.length);

                        label.setObjectText('BARCODE', pngBase64);
alert(records[0]["theme"]);

                        label.print(printersSelect.value);
                    }
                    catch(e)
                    {
                        alert(e.message || e);
                    }
                };
                img.onerror = function()
                {
                    alert('Unable to load qr-code image');                    
                };
                img.src = records[i]["BARCODE"];	
	       //img.src = 'https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=http%3A//developers.dymo.com&choe=UTF-8';	
	
			
			
			
			

		
//                    var pngData = label.render();
//
//                    var labelImage = document.getElementById('img' + (i + 1));
//                    labelImage.src = "data:image/png;base64," + pngData;
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
	    printnow();

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
