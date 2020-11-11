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

            script.setAttribute('src', 'https://docs.google.com/spreadsheets/d/1NuAevHWdpsWChNk20iWUYzTw9iJY1vRmnws7LSjGLv0/1/public/values?alt=json-in-script&callback=window._loadSpreadSheetDataCallback');
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
    <Orientation>Landscape</Orientation>\
    <LabelName>Address30251</LabelName>\
    <InitialLength>0</InitialLength>\
    <BorderStyle>SolidLine</BorderStyle>\
    <DYMORect>\
      <DYMOPoint>\
        <X>0.23</X>\
        <Y>0.06</Y>\
      </DYMOPoint>\
      <Size>\
        <Width>3.21</Width>\
        <Height>0.9966667</Height>\
      </Size>\
    </DYMORect>\
    <BorderColor>\
      <SolidColorBrush>\
        <Color A="1" R="0" G="0" B="0"></Color>\
      </SolidColorBrush>\
    </BorderColor>\
    <BorderThickness>1</BorderThickness>\
    <Show_Border>False</Show_Border>\
    <DynamicLayoutManager>\
      <RotationBehavior>ClearObjects</RotationBehavior>\
      <LabelObjects>\
        <TextObject>\
          <Name>ITextObject1</Name>\
          <Brushes>\
            <BackgroundBrush>\
              <SolidColorBrush>\
                <Color A="0" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </BackgroundBrush>\
            <BorderBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </BorderBrush>\
            <StrokeBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </StrokeBrush>\
            <FillBrush>\
              <SolidColorBrush>\
                <Color A="0" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </FillBrush>\
          </Brushes>\
          <Rotation>Rotation0</Rotation>\
          <OutlineThickness>1</OutlineThickness>\
          <IsOutlined>False</IsOutlined>\
          <BorderStyle>SolidLine</BorderStyle>\
          <Margin>\
            <DYMOThickness Left="0" Top="0" Right="0" Bottom="0" />\
          </Margin>\
          <HorizontalAlignment>Center</HorizontalAlignment>\
          <VerticalAlignment>Middle</VerticalAlignment>\
          <FitMode>AlwaysFit</FitMode>\
          <IsVertical>False</IsVertical>\
          <FormattedText>\
            <FitMode>AlwaysFit</FitMode>\
            <HorizontalAlignment>Center</HorizontalAlignment>\
            <VerticalAlignment>Middle</VerticalAlignment>\
            <IsVertical>False</IsVertical>\
            <LineTextSpan>\
              <TextSpan>\
                <Text> </Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>10.2</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <DataMappingTextSpan>\
                <Text>Beveled Edge: 100mm</Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>10.2</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
                <ColumnName>Column2</ColumnName>\
              </DataMappingTextSpan>\
              <TextSpan>\
                <Text />\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>10.2</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <TextSpan>\
                <Text> </Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>10.2</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <DataMappingTextSpan>\
                <Text>Asian Garden</Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>10.2</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
                <ColumnName>Column3</ColumnName>\
              </DataMappingTextSpan>\
            </LineTextSpan>\
          </FormattedText>\
          <ObjectLayout>\
            <DYMOPoint>\
              <X>0.8819626</X>\
              <Y>0.07</Y>\
            </DYMOPoint>\
            <Size>\
              <Width>1.69102</Width>\
              <Height>0.4983333</Height>\
            </Size>\
          </ObjectLayout>\
        </TextObject>\
        <TextObject>\
          <Name>ITextObject0</Name>\
          <Brushes>\
            <BackgroundBrush>\
              <SolidColorBrush>\
                <Color A="0" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </BackgroundBrush>\
            <BorderBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </BorderBrush>\
            <StrokeBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </StrokeBrush>\
            <FillBrush>\
              <SolidColorBrush>\
                <Color A="0" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </FillBrush>\
          </Brushes>\
          <Rotation>Rotation0</Rotation>\
          <OutlineThickness>1</OutlineThickness>\
          <IsOutlined>False</IsOutlined>\
          <BorderStyle>SolidLine</BorderStyle>\
          <Margin>\
            <DYMOThickness Left="0" Top="0" Right="0" Bottom="0" />\
          </Margin>\
          <HorizontalAlignment>Center</HorizontalAlignment>\
          <VerticalAlignment>Middle</VerticalAlignment>\
          <FitMode>AlwaysFit</FitMode>\
          <IsVertical>False</IsVertical>\
          <FormattedText>\
            <FitMode>AlwaysFit</FitMode>\
            <HorizontalAlignment>Center</HorizontalAlignment>\
            <VerticalAlignment>Middle</VerticalAlignment>\
            <IsVertical>False</IsVertical>\
            <LineTextSpan>\
              <TextSpan>\
                <Text />\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>19.8</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <TextSpan>\
                <Text> </Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>19.8</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <TextSpan>\
                <Text> </Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>19.8</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
              </TextSpan>\
              <DataMappingTextSpan>\
                <Text>(BAG1001)</Text>\
                <FontInfo>\
                  <FontName>Segoe UI</FontName>\
                  <FontSize>19.8</FontSize>\
                  <IsBold>False</IsBold>\
                  <IsItalic>False</IsItalic>\
                  <IsUnderline>False</IsUnderline>\
                  <FontBrush>\
                    <SolidColorBrush>\
                      <Color A="1" R="0" G="0" B="0"></Color>\
                    </SolidColorBrush>\
                  </FontBrush>\
                </FontInfo>\
                <ColumnName>Column4</ColumnName>\
              </DataMappingTextSpan>\
            </LineTextSpan>\
          </FormattedText>\
          <ObjectLayout>\
            <DYMOPoint>\
              <X>0.8819626</X>\
              <Y>0.2958334</Y>\
            </DYMOPoint>\
            <Size>\
              <Width>1.605</Width>\
              <Height>0.4983333</Height>\
            </Size>\
          </ObjectLayout>\
        </TextObject>\
        <BarcodeObject>\
          <Name>IBarcodeObject0</Name>\
          <Brushes>\
            <BackgroundBrush>\
              <SolidColorBrush>\
                <Color A="1" R="1" G="1" B="1"></Color>\
              </SolidColorBrush>\
            </BackgroundBrush>\
            <BorderBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </BorderBrush>\
            <StrokeBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </StrokeBrush>\
            <FillBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </FillBrush>\
          </Brushes>\
          <Rotation>Rotation0</Rotation>\
          <OutlineThickness>1</OutlineThickness>\
          <IsOutlined>False</IsOutlined>\
          <BorderStyle>SolidLine</BorderStyle>\
          <Margin>\
            <DYMOThickness Left="0" Top="0" Right="0" Bottom="0" />\
          </Margin>\
          <BarcodeFormat>Code128Ean</BarcodeFormat>\
          <Data>\
            <MultiDataString>\
              <MappableDataString>Column1</MappableDataString>\
              <DataString></DataString>\
            </MultiDataString>\
          </Data>\
          <HorizontalAlignment>Center</HorizontalAlignment>\
          <VerticalAlignment>Middle</VerticalAlignment>\
          <Size>Small</Size>\
          <TextPosition>Bottom</TextPosition>\
          <FontInfo>\
            <FontName>Arial</FontName>\
            <FontSize>10</FontSize>\
            <IsBold>False</IsBold>\
            <IsItalic>False</IsItalic>\
            <IsUnderline>False</IsUnderline>\
            <FontBrush>\
              <SolidColorBrush>\
                <Color A="1" R="0" G="0" B="0"></Color>\
              </SolidColorBrush>\
            </FontBrush>\
          </FontInfo>\
          <ObjectLayout>\
            <DYMOPoint>\
              <X>0.23</X>\
              <Y>0.4356184</Y>\
            </DYMOPoint>\
            <Size>\
              <Width>0.7634423</Width>\
              <Height>0.4875812</Height>\
            </Size>\
          </ObjectLayout>\
        </BarcodeObject>\
      </LabelObjects>\
    </DynamicLayoutManager>\
  </DYMOLabel>\
  <LabelApplication>Blank</LabelApplication>\
  <DataTable>\
    <Columns>\
      <DataColumn>Column1</DataColumn>\
      <DataColumn>Column2</DataColumn>\
      <DataColumn>Column3</DataColumn>\
      <DataColumn>Column4</DataColumn>\
      <DataColumn>Column5</DataColumn>\
    </Columns>\
    <Rows>\
      <DataRow>\
        <Value>70064649615</Value>\
        <Value>Beveled Edge: 100mm</Value>\
        <Value>Asian Garden</Value>\
        <Value>(BAG1001)</Value>\
        <Value>www.SecretWeaponMiniatures.com</Value>\
      </DataRow>\
    </Rows>\
  </DataTable>\
</DieCutLabel>;
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
