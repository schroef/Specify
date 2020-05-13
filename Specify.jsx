﻿/**
 * Specify
 * =================================
 * Version: 1.3.1
 * https://github.com/adamdehaven/Specify
 *
 * Adam DeHaven
 * http://adamdehaven.com
 * @adamdehaven
 *
 * Additional info:
 * https://adamdehaven.com/blog/dimension-adobe-illustrator-designs-with-a-simple-script/
 * ====================
 */

// Import colorPicker script
 var ScriptFilePath = Folder($.fileName).parent.fsName ; 
$.evalFile(new File(ScriptFilePath + '/colorPicker.inc'));

if (app.documents.length > 0) {

    // Document
    var doc = activeDocument;
    // Count selected items
    var selectedItems = parseInt(doc.selection.length, 10) || 0;

    //
    // Defaults
    // ===========================
    // Units
    var setUnits = true;
    var defaultUnits = $.getenv("Specify_defaultUnits") ? convertToBoolean($.getenv("Specify_defaultUnits")) : setUnits;
    // Font Size
    var setFontSize = 8;
    var defaultFontSize = $.getenv("Specify_defaultFontSize") ? convertToUnits($.getenv("Specify_defaultFontSize")).toFixed(3) : setFontSize;
    // Colors
    var setRed = 36;
    var defaultColorRed = $.getenv("Specify_defaultColorRed") ? $.getenv("Specify_defaultColorRed") : setRed;
    var setGreen = 151;
    var defaultColorGreen = $.getenv("Specify_defaultColorGreen") ? $.getenv("Specify_defaultColorGreen") : setGreen;
    var setBlue = 227;
    var defaultColorBlue = $.getenv("Specify_defaultColorBlue") ? $.getenv("Specify_defaultColorBlue") : setBlue;
    // Head & Tail
    var setHead = 6;
    var defaultHeadTail = $.getenv("Specify_defaultHeadTail") ? $.getenv("Specify_defaultHeadTail") : setHead;
    // Head & Tail
    var setGapSize = 4;
    var defaultGapSize = $.getenv("Specify_defaultGapSize") ? $.getenv("Specify_defaultGapSize") : setGapSize;
    // Stroke width
    var setStroke = 0.5;
    var defaultStroke = $.getenv("Specify_defaultStroke") ? $.getenv("Specify_defaultStroke") : setStroke;
    // Decimals
    var setDecimals = 2;
    var defaultDecimals = $.getenv("Specify_defaultDecimals") ? $.getenv("Specify_defaultDecimals") : setDecimals;
    // Scale
    var setScale = 0;
    var defaultScale = $.getenv("Specify_defaultScale") ? $.getenv("Specify_defaultScale") : setScale;
    // Use Custom Units
    var setUseCustomUnits = false;
    var defaultUseCustomUnits = $.getenv("Specify_defaultUseCustomUnits") ? convertToBoolean($.getenv("Specify_defaultUseCustomUnits")) : setUseCustomUnits;
    // Custom Units
    var setCustomUnits = getRulerUnits();
    var defaultCustomUnits = $.getenv("Specify_defaultCustomUnits") ? $.getenv("Specify_defaultCustomUnits") : setCustomUnits;

    //
    // Create Dialog
    // ===========================

    // Dialog Box
    var specifyDialogBox = new Window("dialog", "Specify");
    specifyDialogBox.alignChildren = "left";
    specifyDialogBox.margins = [20, 20, 20, 20]; // [left, top, right, bottom]

    //
    // Custom Scale Panel
    // ===========================
    scalePanel = specifyDialogBox.add("panel", undefined, "SCALE");
    scalePanel.orientation = "column";
    scalePanel.alignment = "fill";
    scalePanel.margins = 20;
    scalePanel.alignChildren = "left";
    customScaleInfo = scalePanel.add("statictext", undefined, "Define the scale of the artwork/document.");
    customScaleInfo2 = scalePanel.add("statictext", undefined, "Example: 250 units at 1/4 scale displays as 1000");

    // Scale multiplier box
    customScaleGroup = scalePanel.add("group");
    customScaleGroup.orientation = "row";
    customScaleLabel = customScaleGroup.add("statictext", undefined, "Scale:");
    (customScaleDropdown = customScaleGroup.add("dropdownlist")).helpTip = "Choose the scale of the artwork/document. \n\nExample: Choosing '1/4' will indicate the artwork is drawn at \none-fourth scale, resulting in dimension values that are 4x their \ndrawn dimensions.\n\nDefault: 1/1";
    for (var n = 1; n <= 30; n++) {
        if (n == 1) {
            customScaleDropdown.add("item", "1/" + n + "    (Default)");
            customScaleDropdown.add("separator");
        } else {
            customScaleDropdown.add("item", "1/" + n);
        }
    }
    customScaleDropdown.selection = defaultScale;
    customScaleDropdown.onChange = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };

    //
    // Dimension Panel
    // ===========================
    dimensionPanel = specifyDialogBox.add("panel", undefined, "SELECT DIMENSION(S) TO SPECIFY");
    dimensionPanel.orientation = "column";
    dimensionPanel.alignment = "fill";
    dimensionPanel.margins = [20, 20, 20, 10];
    dimensionGroup = dimensionPanel.add("group");
    dimensionGroup.orientation = "row";

    // Top
    (topCheckbox = dimensionGroup.add("checkbox", undefined, "Top")).helpTip = "Dimension the top side of the object(s).";
    topCheckbox.value = false;

    // Right
    (rightCheckbox = dimensionGroup.add("checkbox", undefined, "Right")).helpTip = "Dimension the right side of the object(s).";
    rightCheckbox.value = false;

    // Bottom
    (bottomCheckbox = dimensionGroup.add("checkbox", undefined, "Bottom")).helpTip = "Dimension the bottom side of the object(s).";
    bottomCheckbox.value = false;

    // Left
    (leftCheckbox = dimensionGroup.add("checkbox", undefined, "Left")).helpTip = "Dimension the left side of the object(s).";
    leftCheckbox.value = false;

    // Select All
    selectAllGroup = dimensionPanel.add("group");
    selectAllGroup.orientation = "row";

    (selectAllCheckbox = selectAllGroup.add("checkbox", undefined, "Select All")).helpTip = "Dimension all sides of the object(s).";
    selectAllCheckbox.value = false;
    selectAllCheckbox.onClick = function() {
        if (selectAllCheckbox.value) {
            // Select All is checked
            topCheckbox.value = true;
            topCheckbox.enabled = false;

            rightCheckbox.value = true;
            rightCheckbox.enabled = false;

            bottomCheckbox.value = true;
            bottomCheckbox.enabled = false;

            leftCheckbox.value = true;
            leftCheckbox.enabled = false;
        } else {
            // Select All is unchecked
            topCheckbox.value = false;
            topCheckbox.enabled = true;

            rightCheckbox.value = false;
            rightCheckbox.enabled = true;

            bottomCheckbox.value = false;
            bottomCheckbox.enabled = true;

            leftCheckbox.value = false;
            leftCheckbox.enabled = true;
        }
    };

    // If exactly 2 objects are selected, give user option to dimension BETWEEN them
    if (selectedItems == 2) {
        (between = dimensionPanel.add("checkbox", undefined, "Dimension between selected")).helpTip = "When checked, return the distance between\nthe 2 objects for the selected dimensions.";
        between.value = false;
    }

    //
    // Options Panel
    // ===========================
    optionsPanel = specifyDialogBox.add("panel", undefined, "OPTIONS");
    optionsPanel.orientation = "column";
    optionsPanel.alignment = "fill";
    optionsPanel.margins = 20;
    optionsPanel.alignChildren = "left";

    // Add font-size box
    fontGroup = optionsPanel.add("group");
    fontGroup.orientation = "row";
    fontLabel = fontGroup.add("statictext", undefined, "Font size:");
    (fontSizeInput = fontGroup.add("edittext", undefined, defaultFontSize)).helpTip = "Enter the desired font size for the dimension label(s). If value is less than one (e.g. 0.25) you must include a leading zero before the decimal point.\n\nDefault: " + setFontSize;
    fontUnitsLabelText = "";
    switch (doc.rulerUnits) {
        case RulerUnits.Picas:
            fontUnitsLabelText = "pc";
            break;
        case RulerUnits.Inches:
            fontUnitsLabelText = "in";
            break;
        case RulerUnits.Millimeters:
            fontUnitsLabelText = "mm";
            break;
        case RulerUnits.Centimeters:
            fontUnitsLabelText = "cm";
            break;
        case RulerUnits.Pixels:
            fontUnitsLabelText = "px";
            break;
        default:
            fontUnitsLabelText = "pt";
    }
    fontUnitsLabel = fontGroup.add("statictext", undefined, fontUnitsLabelText);
    fontSizeInput.characters = 5;
    fontSizeInput.onActivate = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    }
    fontSizeInput.onDeactivate = function() {
        // If first character is decimal point, don't error, but instead
        // add leading zero to string.
        if (fontSizeInput.text.charAt(0) == ".") {
            fontSizeInput.text = "0" + fontSizeInput.text;
            fontSizeInput.active = true;
        }
    }
    // Add Color Group
    colorGroup = optionsPanel.add("group");
    colorGroup.orientation = "row";
    colorLabel = colorGroup.add("statictext", undefined, "Label color:");

    // Color Pickr
    function customDraw() {
        with(this) {
            graphics.drawOSControl();
            graphics.rectPath(0, 0, size[0], size[1]);
            graphics.fillPath(fillBrush);
            if (text) graphics.drawString(text, textPen, (size[0] - graphics.measureString(text, graphics.font, size[0])[0]) / 2, 3, graphics.font);
        }
    }

    var colorBtn = colorGroup.add('iconbutton', undefined, undefined, {
        name: 'colorBtn',
        style: 'toolbutton'
    });
    colorBtn.size = [135, 20];
    var colRGB = [(parseInt(defaultColorRed)/255), (parseInt(defaultColorGreen)/255), (parseInt(defaultColorBlue)/255)];
    colorBtn.fillBrush = colorBtn.graphics.newBrush(colorBtn.graphics.BrushType.SOLID_COLOR, colRGB, 1);
    colorBtn.text = "click";
    colorBtn.textPen = colorBtn.graphics.newPen(colorBtn.graphics.PenType.SOLID_COLOR, [1, 1, 1], 1);
    colorBtn.onDraw = customDraw;
    colorBtn.helpTip = "Set a color to use for dimension label(s).";

    var colorInputRed = defaultColorRed;
    var colorInputGreen = defaultColorGreen;
    var colorInputBlue = defaultColorBlue;
    var resultColor = colRGB.toString();
    var colorSet = false;
    function setColor(resultColor) {
        var Red = resultColor[0] * 255;
        var Green = resultColor[1] * 255;
        var Blue = resultColor[2] * 255;
        colorInputRed = Red.toString();
        colorInputGreen = Green.toString();
        colorInputBlue = Blue.toString();
        return [colorInputRed,colorInputGreen,colorInputBlue] 
    }
    colorBtn.onClick = function() {
        colorSet = true;
        resultColor = colorPicker();
        colorBtn.fillBrush = colorBtn.graphics.newBrush(colorBtn.graphics.BrushType.SOLID_COLOR, resultColor);
        colorBtn.text = "";
        colorBtn.onDraw = customDraw;
        // specifyDialogBox.update();
        updatePanel(specifyDialogBox);
        // app.refresh();
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    }
    
    function updatePanel(win) {
        specifyDialogBox.layout.layout(true);
    }

    colorInputRed.onActivate = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    colorInputRed.onChange = function() {
        colorInputRed.text = colorInputRed.text.replace(/[^0-9]/g, "");
    };

    colorInputGreen.onActivate = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    colorInputGreen.onChange = function() {
        colorInputGreen.text = colorInputGreen.text.replace(/[^0-9]/g, "");
    };

    colorInputBlue.onActivate = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    colorInputBlue.onChange = function() {
        colorInputBlue.text = colorInputBlue.text.replace(/[^0-9]/g, "");
    };

    // Size measurement head
    headTailGroup = optionsPanel.add("group");
    headTailGroup.orientation = "row";
    headTailLabel = headTailGroup.add("statictext", undefined, "Head & Tail:");
    (headTailInput = headTailGroup.add("edittext", undefined, defaultHeadTail)).helpTip = "Set size of measurement head & tail.\n\nDefault: " + setHead;
    headTailInput.characters = 3;

    headTailInput.onChange = function() {
        headTailInput.text = headTailInput.text.replace(/[^0-9]/g, "");
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    
    // Gap between measurement lines and object
    gapLabel = headTailGroup.add("statictext", undefined, "Gap:");
    (gapSizeInput = headTailGroup.add("edittext", undefined, defaultGapSize)).helpTip = "Set gap size between measurement lines and object\n\nDefault: " + setGapSize;
    gapSizeInput.characters = 3;

    gapSizeInput.onChange = function() {
        gapSizeInput.text = gapSizeInput.text.replace(/[^0-9.]/g, "");
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    gapSizeInput.onDeactivate = function() {
        // If first character is decimal point, don't error, but instead
        // add leading zero to string.
        if (gapSizeInput.text.charAt(0) == ".") {
            gapSizeInput.text = "0" + gapSizeInput.text;
            gapSizeInput.active = true;
        }
    }

    // Stroke Width
    strokeLabel = headTailGroup.add("statictext", undefined, "Stroke:");
    (strokeInput = headTailGroup.add("edittext", undefined, defaultStroke)).helpTip = "Set width of stroke measurement lines\n\nDefault: " + setStroke;
    strokeInput.characters = 3;

    strokeInput.onChange = function() {
        strokeInput.text = strokeInput.text.replace(/[^0-9.]/g, "");
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    strokeInput.onDeactivate = function() {
        // If first character is decimal point, don't error, but instead
        // add leading zero to string.
        if (strokeInput.text.charAt(0) == ".") {
            strokeInput.text = "0" + strokeInput.text;
            strokeInput.active = true;
        }
    }

    // Add decimal places box
    decimalPlacesGroup = optionsPanel.add("group");
    decimalPlacesGroup.orientation = "row";
    decimalPlacesLabel = decimalPlacesGroup.add("statictext", undefined, "Decimals:"); // Num. of decimal places displayed:
    (decimalPlacesInput = decimalPlacesGroup.add("edittext", undefined, defaultDecimals)).helpTip = "Enter the desired number of decimal places to\ndisplay in the label dimensions.\n\nDefault: " + setDecimals;
    decimalPlacesInput.characters = 4;
    decimalPlacesInput.onActivate = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    decimalPlacesInput.onChange = function() {
        decimalPlacesInput.text = decimalPlacesInput.text.replace(/[^0-9]/g, "");
    };

    // Show/hide units
    (units = optionsPanel.add("checkbox", undefined, "Include units label(s)")).helpTip = "When checked, inserts the units label alongside\nthe outputted dimension.\nExample: 220 px";
    units.value = defaultUnits;
    units.onClick = function() {
        if (units.value == false) {
            useCustomUnits.value = false;
            customUnitsInput.text = getRulerUnits();
            customUnitsInput.enabled = false;
        }
    };

    // Custom Units box
    customUnitsGroup = optionsPanel.add("group");
    customUnitsGroup.orientation = "row";

    // Add options panel checkboxes
    (useCustomUnits = customUnitsGroup.add("checkbox", undefined, "Customize Units Label")).helpTip = "When checked, allows user to customize\nthe text of the units label.\nExample: ft";
    useCustomUnits.value = defaultUseCustomUnits;
    useCustomUnits.onClick = function() {
        if (useCustomUnits.value == true) {
            customUnitsInput.enabled = true;
        } else {
            customUnitsInput.text = getRulerUnits();
            customUnitsInput.enabled = false;
        }
    };

    // Custom Units box
    // customUnitsGroup = optionsPanel.add("group");
    // customUnitsGroup.orientation = "row";
    // customUnitsLabel = customUnitsGroup.add("statictext", undefined, "Custom Units Label:");
    (customUnitsInput = customUnitsGroup.add("edittext", undefined, defaultCustomUnits)).helpTip = "Enter the string to display after the dimension \nnumber when using a custom scale.\n\nDefault: " + setCustomUnits;
    customUnitsInput.enabled = defaultUseCustomUnits;
    customUnitsInput.characters = 8;
    customUnitsInput.onChange = function() {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
        customUnitsInput.text = customUnitsInput.text.replace(/[^ a-zA-Z]/g, "");
    };

    //
    // Restore Defaults Group
    // ===========================
    defaultsPanel = specifyDialogBox.add("panel", undefined, "RESTORE DEFAULTS");
    defaultsPanel.orientation = "column";
    defaultsPanel.alignment = "fill";
    defaultsPanel.margins = 20;
    defaultsPanel.alignChildren = "left";

    // Info text
    infoText = defaultsPanel.add("statictext", undefined, "Options are persistent until application is closed.");
    infoText.margins = 20;
    // Disable to make text appear subtle
    infoText.enabled = false;

    // Reset options button
    restoreDefaultsButton = defaultsPanel.add("button", undefined, "Restore Defaults");
    restoreDefaultsButton.alignment = "left";
    restoreDefaultsButton.enabled = (setFontSize != defaultFontSize || setRed != defaultColorRed || setGreen != defaultColorGreen || setBlue != defaultColorBlue || setHead != defaultHeadTail || setGapSize != defaultGapSize || setStroke != defaultStroke || setDecimals != defaultDecimals || setScale != defaultScale || setCustomUnits != defaultCustomUnits ? true : false);
    restoreDefaultsButton.onClick = function() {
        restoreDefaults();
    };

    function restoreDefaults() {
        units.value = setUnits;
        fontSizeInput.text = setFontSize;
        colorInputRed = setRed;
        colorInputGreen = setGreen;
        colorInputBlue = setBlue;
        headTailInput.text = setHead;
        gapSizeInput.text = setGapSize;
        strokeInput.text = setStroke;
        decimalPlacesInput.text = setDecimals;
        customScaleDropdown.selection = setScale;
        customUnitsInput.text = setCustomUnits;
        useCustomUnits.value = false;
        customUnitsInput.enabled = false;
        restoreDefaultsButton.enabled = false;
        
        // Unset environmental variables
        $.setenv("Specify_defaultUnits", "");
        $.setenv("Specify_defaultFontSize", "");
        $.setenv("Specify_defaultColorRed", "");
        $.setenv("Specify_defaultColorGreen", "");
        $.setenv("Specify_defaultColorBlue", "");
        $.setenv("Specify_defaultHeadTail", "");
        $.setenv("Specify_defaultGapSize", "");
        $.setenv("Specify_defaultStroke", "");
        $.setenv("Specify_defaultDecimals", "");
        $.setenv("Specify_defaultScale", "");
        $.setenv("Specify_defaultUseCustomUnits", "");
        $.setenv("Specify_defaultCustomUnits", "");

        colorBtn.text = "click";
        colorBtn.fillBrush = colorBtn.graphics.newBrush(colorBtn.graphics.BrushType.SOLID_COLOR, [setRed,setGreen,setBlue], 1);
        colorBtn.onDraw = customDraw;
        // updatePanel(specifyDialogBox); // Force redraw dialog
    };

    //
    // Button Group
    // ===========================
    buttonGroup = specifyDialogBox.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";
    buttonGroup.margins = [20, 0, 0, 0]; // [left, top, right, bottom]

    // Cancel button
    cancelButton = buttonGroup.add("button", undefined, "Cancel");
    cancelButton.onClick = function() {
        specifyDialogBox.close();
    };

    // Specify button
    specifyButton = buttonGroup.add("button", undefined, "Specify Object(s)");
    specifyButton.size = [125, 40];
    specifyDialogBox.defaultElement = specifyButton;
    specifyButton.onClick = function() {
        startSpec();
    };

    //
    // Footer Group
    // ===========================

    // footerGroup = specifyDialogBox.add("group");
    // footerGroup.orientation = "row";
    // footerGroup.alignment = "center";
    // footerGroup.margins = [10, 0, 0, 0]; // [left, top, right, bottom]

    // // Updates & Info
    // specifyUpdate = footerGroup.add("statictext", undefined, "https://github.com/adamdehaven/Specify"); // Updates & Info:  removed makes it less wide 

    //
    // ===========================
    // End Create Dialog

    //
    // SPEC Layer
    // ===========================
    try {
        var specsLayer = doc.layers["SPECS"];
    } catch (err) {
        var specsLayer = doc.layers.add();
        specsLayer.name = "SPECS";
    }

    // Measurement line and text color in RGB
    var color = new RGBColor;

    // Declare global decimals var
    var decimals;

    // Declare global scale var
    var scale;

    // // Gap between measurement lines and object
    // var gap = 4;

    // Size of perpendicular measurement lines.
    // var size = 6;
    // var size = parseInt(headTailInput.text);
    // alert("size " +size)

    //
    // Start the Spec
    // ===========================
    function startSpec() {

        // Add all selected objects to array
        var objectsToSpec = new Array();
        for (var index = doc.selection.length - 1; index >= 0; index--) {
            objectsToSpec[index] = doc.selection[index];
        }

        // Fetch desired dimensions
        var top = topCheckbox.value;
        var left = leftCheckbox.value;
        var right = rightCheckbox.value;
        var bottom = bottomCheckbox.value;
        // Take focus away from fontSizeInput to validate (numeric)
        fontSizeInput.active = false;

        // Set bool for numeric vars
        var validFontSize = /^[0-9]{1,3}(\.[0-9]{1,3})?$/.test(fontSizeInput.text);

        // var validRedColor = /^[0-9]{1,3}$/.test(colorInputRed.text) && parseInt(colorInputRed.text) >= 0 && parseInt(colorInputRed.text) <= 255;
        // var validGreenColor = /^[0-9]{1,3}$/.test(colorInputGreen.text) && parseInt(colorInputGreen.text) >= 0 && parseInt(colorInputRed.text) <= 255;
        // var validBlueColor = /^[0-9]{1,3}$/.test(colorInputBlue.text) && parseInt(colorInputBlue.text) >= 0 && parseInt(colorInputBlue.text) <= 255;
        // var rgbCol = setColor(resultColor);
        // alert(setColor(resultColor))
        // alert(setColor(resultColor)[1])
        // alert(colorSet)
        if (colorSet) {
            colorInputRed = setColor(resultColor)[0];
            colorInputGreen = setColor(resultColor)[1];
            colorInputBlue = setColor(resultColor)[2];
        }
        var validRedColor = /^[0-9]{1,3}$/.test(colorInputRed) && colorInputRed >= 0 && colorInputRed <= 255;
        var validGreenColor = /^[0-9]{1,3}$/.test(colorInputGreen) && colorInputGreen >= 0 && colorInputRed <= 255;
        var validBlueColor = /^[0-9]{1,3}$/.test(colorInputBlue) && colorInputBlue >= 0 && colorInputBlue <= 255;
        // If colors are valid, set variables
        if (validRedColor && validGreenColor && validBlueColor) {
            // color.red = colorInputRed.text;
            // color.green = colorInputGreen.text;
            // color.blue = colorInputBlue.text;
            color.red = colorInputRed;
            color.green = colorInputGreen;
            color.blue = colorInputBlue;
            // Set environmental variables
            $.setenv("Specify_defaultColorRed", color.red);
            $.setenv("Specify_defaultColorGreen", color.green);
            $.setenv("Specify_defaultColorBlue", color.blue);
            colorSet = false;
        }

        var validHead = /^[0-9]{1,2}$/.test(headTailInput.text);
        if (validHead) {
            size = parseInt(headTailInput.text);
            $.setenv("Specify_defaultHeadTail", size);
        }
        
        var validGap = /^[0-9]{1,3}(\.[0-9]{1,3})?$/.test(gapSizeInput.text);
        if (validGap) {
            gap = parseInt(gapSizeInput.text);
            $.setenv("Specify_defaultGapSize", gap);
        }
        
        var validWidth = /^[0-9]{1,3}(\.[0-9]{1,3})?$/.test(strokeInput.text);
        if (validWidth) {
            strWidth = parseFloat(strokeInput.text);
            $.setenv("Specify_defaultStroke", strWidth);
        }

        var validDecimalPlaces = /^[0-4]{1}$/.test(decimalPlacesInput.text);
        if (validDecimalPlaces) {
            // Number of decimal places in measurement
            decimals = decimalPlacesInput.text;
            // Set environmental variable
            $.setenv("Specify_defaultDecimals", decimals);
        }

        var theScale = parseInt(customScaleDropdown.selection.toString().replace(/1\//g, "").replace(/[^0-9]/g, ""));
        scale = theScale;
        // Set environmental variable
        $.setenv("Specify_defaultScale", customScaleDropdown.selection.index);

        if (selectedItems < 1) {
            beep();
            alert("Please select at least 1 object and try again.");
            // Close dialog
            specifyDialogBox.close();
        } else if (!top && !left && !right && !bottom) {
            beep();
            alert("Please select at least 1 dimension to draw.");
        } else if (!validFontSize) {
            // If fontSizeInput.text does not match regex
            beep();
            alert("Please enter a valid font size. \n0.002 - 999.999");
            fontSizeInput.active = true;
            fontSizeInput.text = setFontSize;
        } else if (parseFloat(fontSizeInput.text, 10) <= 0.001) {
            beep();
            alert("Font size must be greater than 0.001.");
            fontSizeInput.active = true;
        } else if (!validRedColor || !validGreenColor || !validBlueColor) {
            // If RGB inputs are not numeric
            beep();
            alert("Please enter a valid RGB color.");
            colorInputRed.active = true;
            colorInputRed.text = setRed;
            colorInputGreen.text = setGreen;
            colorInputBlue.text = setBlue;
        } else if (!validHead) {
            // If inputs are not numeric
            beep();
            alert("Please enter a valid number for head & tail.");
            headTailInput.active = true;
            headTailInput.text = setHead;
        } else if (!validGap) {
            // If inputs are not numeric
            beep();
            alert("Please enter a valid number for gap size.");
            gapSizeInput.active = true;
            gapSizeInput.text = setGapSize;
        } else if (!validWidth) {
            // If inputs are not numeric
            beep();
            alert("Please enter a valid number for strokeInput.");
            strokeInput.active = true;
            strokeInput.text = setStroke;
        }  else if (!validDecimalPlaces) {
            // If decimalPlacesInput.text is not numeric
            beep();
            alert("Decimal places must range from 0 - 4.");
            decimalPlacesInput.active = true;
            decimalPlacesInput.text = setDecimals;
        } else if (selectedItems == 2 && between.value) {
            if (top) specDouble(objectsToSpec[0], objectsToSpec[1], "Top");
            if (left) specDouble(objectsToSpec[0], objectsToSpec[1], "Left");
            if (right) specDouble(objectsToSpec[0], objectsToSpec[1], "Right");
            if (bottom) specDouble(objectsToSpec[0], objectsToSpec[1], "Bottom");
            // Close dialog when finished
            specifyDialogBox.close();
        } else {
            // Iterate over each selected object, creating individual dimensions as you go
            for (var objIndex = objectsToSpec.length - 1; objIndex >= 0; objIndex--) {
                if (top) specSingle(objectsToSpec[objIndex].geometricBounds, "Top");
                if (left) specSingle(objectsToSpec[objIndex].geometricBounds, "Left");
                if (right) specSingle(objectsToSpec[objIndex].geometricBounds, "Right");
                if (bottom) specSingle(objectsToSpec[objIndex].geometricBounds, "Bottom");
            }
            // Close dialog when finished
            specifyDialogBox.close();
        }
    };

    //
    // Spec a single object
    // ===========================
    function specSingle(bound, where) {
        // unlock SPECS layer
        specsLayer.locked = false;

        // width and height
        var w = bound[2] - bound[0];
        var h = bound[1] - bound[3];

        // a & b are the horizontal or vertical positions that change
        // c is the horizontal or vertical position that doesn't change
        var a = bound[0];
        var b = bound[2];
        var c = bound[1];

        // xy='x' (horizontal measurement), xy='y' (vertical measurement)
        var xy = "x";

        // a direction flag for placing the measurement lines.
        var dir = 1;

        // Convert gap (px) size to RulerUnits
        var gapUnit = convertToGap(gap, "px");
        function convertToGap(value, unit) {
            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    value = new UnitValue(value, "pc").as(unit);
                    break;
                case RulerUnits.Inches:
                    value = new UnitValue(value, "in").as(unit);
                    break;
                case RulerUnits.Millimeters:
                    value = new UnitValue(value, "mm").as(unit);
                    break;
                case RulerUnits.Centimeters:
                    value = new UnitValue(value, "cm").as(unit);
                    break;
                case RulerUnits.Pixels:
                    value = new UnitValue(value, "px").as(unit);
                    break;
                default:
                    value = new UnitValue(value, unit).as(unit);
            }
            return value;
        };

        // Convert size (px) size to RulerUnits
        var sizeUnit = convertToGap(size, "px");
        function convertToGap(value, unit) {
            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    value = new UnitValue(value, "pc").as(unit);
                    break;
                case RulerUnits.Inches:
                    value = new UnitValue(value, "in").as(unit);
                    break;
                case RulerUnits.Millimeters:
                    value = new UnitValue(value, "mm").as(unit);
                    break;
                case RulerUnits.Centimeters:
                    value = new UnitValue(value, "cm").as(unit);
                    break;
                case RulerUnits.Pixels:
                    value = new UnitValue(value, "px").as(unit);
                    break;
                default:
                    value = new UnitValue(value, unit).as(unit);
            }
            return value;
        };
        
        switch (where) {
            case "Top":
                a = bound[0];
                b = bound[2];
                c = bound[1];
                xy = "x";
                dir = 1;
                break;
            case "Right":
                a = bound[1];
                b = bound[3];
                c = bound[2];
                xy = "y";
                dir = 1;
                break;
            case "Bottom":
                a = bound[0];
                b = bound[2];
                c = bound[3];
                xy = "x";
                dir = -1;
                break;
            case "Left":
                a = bound[1];
                b = bound[3];
                c = bound[0];
                xy = "y";
                dir = -1;
                break;
        }

        // Create the measurement lines
        var lines = new Array();

        // horizontal measurement
        if (xy == "x") {

            // 2 vertical lines
            lines[0] = new Array(new Array(a, c + (gapUnit) * dir));
            lines[0].push(new Array(a, c + (gapUnit + sizeUnit) * dir));
            lines[1] = new Array(new Array(b, c + (gapUnit) * dir));
            lines[1].push(new Array(b, c + (gapUnit + sizeUnit) * dir));

            // 1 horizontal line
            lines[2] = new Array(new Array(a, c + (gapUnit + sizeUnit / 2) * dir));
            lines[2].push(new Array(b, c + (gapUnit + sizeUnit / 2) * dir));

            // Create text label
            if (where == "Top") {
                var t = specLabel(w, (a + b) / 2, lines[0][1][1], color);
                t.top += t.height;
            } else {
                var t = specLabel(w, (a + b) / 2, lines[0][0][1], color);
                t.top -= sizeUnit;
            }
            t.left -= t.width / 2;

        } else {
            // Vertical measurement

            // 2 horizontal lines
            lines[0] = new Array(new Array(c + (gapUnit) * dir, a));
            lines[0].push(new Array(c + (gapUnit + sizeUnit) * dir, a));
            lines[1] = new Array(new Array(c + (gapUnit) * dir, b));
            lines[1].push(new Array(c + (gapUnit + sizeUnit) * dir, b));

            //1 vertical line
            lines[2] = new Array(new Array(c + (gapUnit + sizeUnit / 2) * dir, a));
            lines[2].push(new Array(c + (gapUnit + sizeUnit / 2) * dir, b));

            // Create text label
            if (where == "Left") {
                var t = specLabel(h, lines[0][1][0], (a + b) / 2, color);
                t.left -= t.width;
                t.rotate(90, true, false, false, false, Transformation.BOTTOMRIGHT);
                t.top += t.width;
                t.top += t.height / 2;
            } else {
                var t = specLabel(h, lines[0][1][0], (a + b) / 2, color);
                t.rotate(-90, true, false, false, false, Transformation.BOTTOMLEFT);
                t.top += t.width;
                t.top += t.height / 2;
            }
        }

        // Draw lines
        var specgroup = new Array(t);

        for (var i = 0; i < lines.length; i++) {
            var p = doc.pathItems.add();
            p.setEntirePath(lines[i]);
            p.strokeDashes = []; // Prevent dashed SPEC lines
            setLineStyle(p, color, strWidth);
            specgroup.push(p);
        }

        group(specsLayer, specgroup);

        // re-lock SPECS layer
        specsLayer.locked = true;

    };

    //
    // Spec the gap between 2 elements
    // ===========================
    function specDouble(item1, item2, where) {

        var bound = new Array(0, 0, 0, 0);

        var a = item1.geometricBounds;
        var b = item2.geometricBounds;

        if (where == "Top" || where == "Bottom") {

            if (b[0] > a[0]) { // item 2 on right,

                if (b[0] > a[2]) { // no overlap
                    bound[0] = a[2];
                    bound[2] = b[0];
                } else { // overlap
                    bound[0] = b[0];
                    bound[2] = a[2];
                }
            } else if (a[0] >= b[0]) { // item 1 on right

                if (a[0] > b[2]) { // no overlap
                    bound[0] = b[2];
                    bound[2] = a[0];
                } else { // overlap
                    bound[0] = a[0];
                    bound[2] = b[2];
                }
            }

            bound[1] = Math.max(a[1], b[1]);
            bound[3] = Math.min(a[3], b[3]);

        } else {

            if (b[3] > a[3]) { // item 2 on top
                if (b[3] > a[1]) { // no overlap
                    bound[3] = a[1];
                    bound[1] = b[3];
                } else { // overlap
                    bound[3] = b[3];
                    bound[1] = a[1];
                }
            } else if (a[3] >= b[3]) { // item 1 on top

                if (a[3] > b[1]) { // no overlap
                    bound[3] = b[1];
                    bound[1] = a[3];
                } else { // overlap
                    bound[3] = a[3];
                    bound[1] = b[1];
                }
            }

            bound[0] = Math.min(a[0], b[0]);
            bound[2] = Math.max(a[2], b[2]);
        }
        specSingle(bound, where);
    };

    //
    // Create a text label that specify the dimension
    // ===========================
    function specLabel(val, x, y, color) {

        var t = doc.textFrames.add();
        // Get font size from specifyDialogBox.fontSizeInput
        var labelFontSize;
        if (parseFloat(fontSizeInput.text) > 0) {
            labelFontSize = parseFloat(fontSizeInput.text);
        } else {
            labelFontSize = defaultFontSize;
        }

        // Convert font size to RulerUnits
        var labelFontInUnits = convertToPoints(labelFontSize, "pt");

        // Set environmental variable
        $.setenv("Specify_defaultFontSize", labelFontInUnits);

        t.textRange.characterAttributes.size = labelFontInUnits;
        t.textRange.characterAttributes.alignment = StyleRunAlignmentType.center;
        t.textRange.characterAttributes.fillColor = color;

        // Conversions : http://wwwimages.adobe.com/content/dam/Adobe/en/devnet/illustrator/sdk/CC2014/Illustrator%20Scripting%20Guide.pdf
        // UnitValue object (page 230): http://wwwimages.adobe.com/content/dam/Adobe/en/devnet/scripting/pdfs/javascript_tools_guide.pdf

        var displayUnitsLabel = units.value;
        // Set environmental variable
        $.setenv("Specify_defaultUnits", displayUnitsLabel);

        var v = val * scale;
        var unitsLabel = "";

        switch (doc.rulerUnits) {
            case RulerUnits.Picas:
                v = new UnitValue(v, "pt").as("pc");
                var vd = v - Math.floor(v);
                vd = 12 * vd;
                v = Math.floor(v) + "p" + vd.toFixed(decimals);
                break;
            case RulerUnits.Inches:
                v = new UnitValue(v, "pt").as("in");
                v = v.toFixed(decimals);
                unitsLabel = " in"; // add abbreviation
                break;
            case RulerUnits.Millimeters:
                v = new UnitValue(v, "pt").as("mm");
                v = v.toFixed(decimals);
                unitsLabel = " mm"; // add abbreviation
                break;
            case RulerUnits.Centimeters:
                v = new UnitValue(v, "pt").as("cm");
                v = v.toFixed(decimals);
                unitsLabel = " cm"; // add abbreviation
                break;
            case RulerUnits.Pixels:
                v = new UnitValue(v, "pt").as("px");
                v = v.toFixed(decimals);
                unitsLabel = " px"; // add abbreviation
                break;
            default:
                v = new UnitValue(v, "pt").as("pt");
                v = v.toFixed(decimals);
                unitsLabel = " pt"; // add abbreviation
        }

        // If custom scale and units label is set
        if (useCustomUnits.value == true && customUnitsInput.enabled && customUnitsInput.text != getRulerUnits()) {
            unitsLabel = customUnitsInput.text;
            $.setenv("Specify_defaultUseCustomUnits", true);
            $.setenv("Specify_defaultCustomUnits", unitsLabel);
        }

        if (displayUnitsLabel) {
            t.contents = v + " " + unitsLabel;
        } else {
            t.contents = v;
        }
        t.top = y;
        t.left = x;

        return t;
    };

    function convertToBoolean(string) {
        switch (string.toLowerCase()) {
            case "true":
                return true;
                break;
            case "false":
                return false;
                break;
        }
    };

    function setLineStyle(path, color, strWidth) {
        path.filled = false;
        path.stroked = true;
        path.strokeColor = color;
        path.strokeWidth = strWidth;
        return path;
    };

    // Group items in a layer
    function group(layer, items, isDuplicate) {

        // Create new group
        var gg = layer.groupItems.add();

        // Add to group
        // Reverse count, because items length is reduced as items are moved to new group
        for (var i = items.length - 1; i >= 0; i--) {

            if (items[i] != gg) { // don't group the group itself
                if (isDuplicate) {
                    newItem = items[i].duplicate(gg, ElementPlacement.PLACEATBEGINNING);
                } else {
                    items[i].move(gg, ElementPlacement.PLACEATBEGINNING);
                }
            }
        }
        return gg;
    };

    function convertToPoints(value) {
        switch (doc.rulerUnits) {
            case RulerUnits.Picas:
                value = new UnitValue(value, "pc").as("pt");
                break;
            case RulerUnits.Inches:
                value = new UnitValue(value, "in").as("pt");
                break;
            case RulerUnits.Millimeters:
                value = new UnitValue(value, "mm").as("pt");
                break;
            case RulerUnits.Centimeters:
                value = new UnitValue(value, "cm").as("pt");
                break;
            case RulerUnits.Pixels:
                value = new UnitValue(value, "px").as("pt");
                break;
            default:
                value = new UnitValue(value, "pt").as("pt");
        }
        return value;
    };

    function convertToUnits(value) {
        switch (doc.rulerUnits) {
            case RulerUnits.Picas:
                value = new UnitValue(value, "pt").as("pc");
                break;
            case RulerUnits.Inches:
                value = new UnitValue(value, "pt").as("in");
                break;
            case RulerUnits.Millimeters:
                value = new UnitValue(value, "pt").as("mm");
                break;
            case RulerUnits.Centimeters:
                value = new UnitValue(value, "pt").as("cm");
                break;
            case RulerUnits.Pixels:
                value = new UnitValue(value, "pt").as("px");
                break;
            default:
                value = new UnitValue(value, "pt").as("pt");
        }
        return value;
    };

    function getRulerUnits() {
        var rulerUnits;
        switch (doc.rulerUnits) {
            case RulerUnits.Picas:
                rulerUnits = "pc";
                break;
            case RulerUnits.Inches:
                rulerUnits = "in";
                break;
            case RulerUnits.Millimeters:
                rulerUnits = "mm";
                break;
            case RulerUnits.Centimeters:
                rulerUnits = "cm";
                break;
            case RulerUnits.Pixels:
                rulerUnits = "px";
                break;
            default:
                rulerUnits = "pt";
        }
        return rulerUnits;
    };

    //
    // Run Script
    // ===========================
    switch (selectedItems) {
        case 0:
            beep();
            alert("Please select at least 1 object and try again.");
            break;
        default:
            specifyDialogBox.show();
            break;
    }

} else { // No active document
    alert("There are no objects to Specify. \nPlease open a document to continue.")
}