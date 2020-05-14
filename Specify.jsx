/**
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
    // Colors Label
    var setLabelRed = 57; //36;
    var defaultLabelColorRed = $.getenv("Specify_defaultLabelColorRed") ? $.getenv("Specify_defaultLabelColorRed") : setLabelRed;
    var setLabelGreen = 140; //151;
    var defaultLabelColorGreen = $.getenv("Specify_defaultLabelColorGreen") ? $.getenv("Specify_defaultLabelColorGreen") : setLabelGreen;
    var setLabelBlue = 242; //227;
    var defaultLabelColorBlue = $.getenv("Specify_defaultLabelColorBlue") ? $.getenv("Specify_defaultLabelColorBlue") : setLabelBlue;
     // Colors Lines
    var setLinesRed =  57; //36;
    var defaultLinesColorRed = $.getenv("Specify_defaultLinesColorRed") ? $.getenv("Specify_defaultLinesColorRed") : setLinesRed;
    var setLinesGreen = 140; //151;
    var defaultLinesColorGreen = $.getenv("Specify_defaultLinesColorGreen") ? $.getenv("Specify_defaultLinesColorGreen") : setLinesGreen;
    var setLinesBlue = 242; //227;
    var defaultLinesColorBlue = $.getenv("Specify_defaultLinesColorBlue") ? $.getenv("Specify_defaultLinesColorBlue") : setLinesBlue;
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
    scalePanel.margins = 16;
    scalePanel.spacing = 8;
    scalePanel.alignChildren = "left";
    customScaleInfo = scalePanel.add("statictext", undefined, "Define the scale of the artwork/document. Example: 250 units at 1/4 scale displays as 1000", {multiline:true});
    customScaleInfo.preferredSize.width = 250; 
    // customScaleInfo2 = scalePanel.add("statictext", undefined, "Example: 250 units at 1/4 scale displays as 1000");

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
    // Labels Panel
    // ===========================
    labelsPanel = specifyDialogBox.add("panel", undefined, "LABELS / LINES");
    labelsPanel.orientation = "column";
    labelsPanel.alignment = "fill";
    labelsPanel.margins = 16;
    labelsPanel.spacing = 8;
    labelsPanel.alignChildren = "left";

    // Add font-size box
    fontGroup = labelsPanel.add("group");
    fontGroup.orientation = "row";
    fontGroup.spacing = 5;
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

    // Custom background Buttons
    function customDraw() {
        with(this) {
            graphics.drawOSControl();
            graphics.rectPath(0, 0, size[0], size[1]);
            graphics.fillPath(fillBrush);
            if (text) graphics.drawString(text, textPen, (size[0] - graphics.measureString(text, graphics.font, size[0])[0]) / 2, 3, graphics.font);
        }
    }

    // Refresh Panel
    function updatePanel(win) {
        specifyDialogBox.layout.layout(true);
    }

    // Add Color Label Group
    colLabelGroup = fontGroup.add("group");
    colLabelGroup.orientation = "row";
    colLabel = colLabelGroup.add("statictext", undefined, "");
    
    var colLabelBtn = colLabelGroup.add('iconbutton', undefined, undefined, {
        name: 'colLabelBtn',
        style: 'toolbutton'
    });
    colLabelBtn.size = [70, 20];
    var colLabelRGB = [(parseInt(defaultLabelColorRed)/255), (parseInt(defaultLabelColorGreen)/255), (parseInt(defaultLabelColorBlue)/255)];
    colLabelBtn.fillBrush = colLabelBtn.graphics.newBrush(colLabelBtn.graphics.BrushType.SOLID_COLOR, colLabelRGB, 1);
    // colLabelBtn.text = "click";
    colLabelBtn.textPen = colLabelBtn.graphics.newPen(colLabelBtn.graphics.PenType.SOLID_COLOR, [1, 1, 1], 1);
    colLabelBtn.onDraw = customDraw;
    colLabelBtn.helpTip = "Set a color to use for dimension label(s).";

    var colLabelInputRed = defaultLabelColorRed;
    var colLabelInputGreen = defaultLabelColorGreen;
    var colLabelInputBlue = defaultLabelColorBlue;
    var resultLabelColor = colLabelRGB.toString();
    var colLabelSet = false;

    function setLabelColor(resultLabelColor) {
        var Red = Math.round(resultLabelColor[0] * 255);
        var Green = Math.round(resultLabelColor[1] * 255);
        var Blue = Math.round(resultLabelColor[2] * 255);
        colLabelInputRed = Red.toString();
        colLabelInputGreen = Green.toString();
        colLabelInputBlue = Blue.toString();
        return [colLabelInputRed,colLabelInputGreen,colLabelInputBlue] 
    }

    colLabelBtn.onClick = function() {
        colLabelSet = true;
        resultLabelColor = colorPicker(colLabelRGB);
        colLabelBtn.fillBrush = colLabelBtn.graphics.newBrush(colLabelBtn.graphics.BrushType.SOLID_COLOR, resultLabelColor);
        colLabelBtn.text = "";
        colLabelBtn.onDraw = customDraw;
        // specifyDialogBox.update();
        updatePanel(specifyDialogBox);
        // app.refresh();
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    }

    // Add decimal places box
    decimalPlacesGroup = labelsPanel.add("group");
    decimalPlacesGroup.orientation = "row";
    decimalPlacesGroup.spacing = 5;
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
    (units = labelsPanel.add("checkbox", undefined, "Include units label(s)")).helpTip = "When checked, inserts the units label alongside\nthe outputted dimension.\nExample: 220 px";
    units.value = defaultUnits;
    units.onClick = function() {
        if (units.value == false) {
            useCustomUnits.value = false;
            customUnitsInput.text = getRulerUnits();
            customUnitsInput.enabled = false;
        }
    };

    // Custom Units box
    customUnitsGroup = labelsPanel.add("group");
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
    // customUnitsGroup = labelsPanel.add("group");
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
    // Lines Panel
    // ===========================
    // linesPanel = specifyDialogBox.add("panel", undefined, "OPTIONS");
    // linesPanel.orientation = "column";
    // linesPanel.alignment = "fill";
    // linesPanel.margins = 16;
    // linesPanel.spacing = 8;
    // linesPanel.alignChildren = "left";

    // Size measurement head
    headTailGroup = labelsPanel.add("group");
    headTailGroup.orientation = "row";
    headTailGroup.spacing = 5;
    headTailLabel = headTailGroup.add("statictext", undefined, "Head & Tail:");
    (headTailInput = headTailGroup.add("edittext", undefined, defaultHeadTail)).helpTip = "Set size of measurement head & tail.\n\nDefault: " + setHead;
    headTailInput.characters = 3;
    heaedUnitsLabel = headTailGroup.add("statictext", undefined, fontUnitsLabelText);

    headTailInput.onChange = function() {
        headTailInput.text = headTailInput.text.replace(/[^0-9]/g, "");
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    
    // Gap between measurement lines and object
    spacer01 = headTailGroup.add("statictext", undefined, "  ");
    gapLabel = headTailGroup.add("statictext", undefined, "Gap:");
    (gapSizeInput = headTailGroup.add("edittext", undefined, defaultGapSize)).helpTip = "Set gap size between measurement lines and object\n\nDefault: " + setGapSize;
    gapSizeInput.characters = 3;
    gapUnitsLabel = headTailGroup.add("statictext", undefined, fontUnitsLabelText);

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

    // Stroke Group
    strokeGroup = labelsPanel.add("group");
    strokeGroup.orientation = "row";
    strokeGroup.spacing = 5;
    // Stroke Width
    strokeLabel = strokeGroup.add("statictext", undefined, "Stroke:");
    (strokeInput = strokeGroup.add("edittext", undefined, defaultStroke)).helpTip = "Set width of stroke measurement lines\n\nDefault: " + setStroke;
    strokeInput.characters = 3;
    strokeUnitsLabel = strokeGroup.add("statictext", undefined, fontUnitsLabelText);

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

    // Add Color Lines
    colLines = strokeGroup.add("statictext", undefined, "");
    var colLinesBtn = strokeGroup.add('iconbutton', undefined, undefined, {
        name: 'colLinesBtn',
        style: 'toolbutton'
    });
    colLinesBtn.size = [70, 20];
    var colLinesRGB = [(parseInt(defaultLinesColorRed)/255), (parseInt(defaultLinesColorGreen)/255), (parseInt(defaultLinesColorBlue)/255)];
    colLinesBtn.fillBrush = colLinesBtn.graphics.newBrush(colLinesBtn.graphics.BrushType.SOLID_COLOR, colLinesRGB, 1);
    // colLinesBtn.text = "click";
    colLinesBtn.textPen = colLinesBtn.graphics.newPen(colLinesBtn.graphics.PenType.SOLID_COLOR, [1, 1, 1], 1);
    colLinesBtn.onDraw = customDraw;
    colLinesBtn.helpTip = "Set a color to use for dimension Line(s).";

    var colLinesInputRed = defaultLinesColorRed;
    var colLinesInputGreen = defaultLinesColorGreen;
    var colLinesInputBlue = defaultLinesColorBlue;
    var resultLinesColor = colLinesRGB.toString();
    var colLinesSet = false;

    function setLinesColor(resultLinesColor) {
        var Red = Math.round(resultLinesColor[0] * 255);
        var Green = Math.round(resultLinesColor[1] * 255);
        var Blue = Math.round(resultLinesColor[2] * 255);
        colLinesInputRed = Red.toString();
        colLinesInputGreen = Green.toString();
        colLinesInputBlue = Blue.toString();
        return [colLinesInputRed,colLinesInputGreen,colLinesInputBlue] 
    }

    colLinesBtn.onClick = function() {
        colLinesSet = true;
        resultLinesColor = colorPicker(colLinesRGB);
        colLinesBtn.fillBrush = colLinesBtn.graphics.newBrush(colLinesBtn.graphics.BrushType.SOLID_COLOR, resultLinesColor);
        colLinesBtn.text = "";
        colLinesBtn.onDraw = customDraw;
        // specifyDialogBox.update();
        updatePanel(specifyDialogBox);
        // app.refresh();
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    }


    //
    // Restore Defaults Group
    // ===========================
    defaultsPanel = specifyDialogBox.add("panel", undefined, "RESTORE DEFAULTS");
    defaultsPanel.orientation = "column";
    defaultsPanel.alignment = "fill";
    defaultsPanel.margins = 16;
    defaultsPanel.spacing = 8;
    defaultsPanel.alignChildren = "left";

    // Info text
    infoText = defaultsPanel.add("statictext", undefined, "Options are persistent until app quit"); //application is closed.
    infoText.margins = 16;
    infoText.spacing = 8;
    // Disable to make text appear subtle
    infoText.enabled = false;

    // Reset options button
    restoreDefaultsButton = defaultsPanel.add("button", undefined, "Restore Defaults");
    restoreDefaultsButton.alignment = "left";
    restoreDefaultsButton.enabled = (setFontSize != defaultFontSize || setLabelRed != defaultLabelColorRed || setLabelGreen != defaultLabelColorGreen || setLabelBlue != defaultLabelColorBlue || setLinesRed != defaultLinesColorRed || setLinesGreen != defaultLinesColorGreen || setLinesBlue != defaultLinesColorBlue || setHead != defaultHeadTail || setGapSize != defaultGapSize || setStroke != defaultStroke || setDecimals != defaultDecimals || setScale != defaultScale || setCustomUnits != defaultCustomUnits ? true : false);
    restoreDefaultsButton.onClick = function() {
        restoreDefaults();
    };

    function restoreDefaults() {
        units.value = setUnits;
        fontSizeInput.text = setFontSize;
        colLabelInputRed = setLabelRed;
        colLabelInputGreen = setLabelGreen;
        colLabelInputBlue = setLabelBlue;
        colLinesInputRed = setLinesRed;
        colLinesInputGreen = setLinesGreen;
        colLinesInputBlue = setLinesBlue;
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
        $.setenv("Specify_defaultLabelColorRed", "");
        $.setenv("Specify_defaultLabelColorGreen", "");
        $.setenv("Specify_defaultLabelColorBlue", "");
        $.setenv("Specify_defaultLinesColorRed", "");
        $.setenv("Specify_defaultLinesColorGreen", "");
        $.setenv("Specify_defaultLinesColorBlue", "");
        $.setenv("Specify_defaultHeadTail", "");
        $.setenv("Specify_defaultGapSize", "");
        $.setenv("Specify_defaultStroke", "");
        $.setenv("Specify_defaultDecimals", "");
        $.setenv("Specify_defaultScale", "");
        $.setenv("Specify_defaultUseCustomUnits", "");
        $.setenv("Specify_defaultCustomUnits", "");

        // colLabelBtn.text = "click";
        colLabelBtn.fillBrush = colLabelBtn.graphics.newBrush(colLabelBtn.graphics.BrushType.SOLID_COLOR, [setLabelRed,setLabelGreen,setLabelBlue], 1);
        colLabelBtn.onDraw = customDraw;
        // colLinesBtn.text = "click";
        colLinesBtn.fillBrush = colLinesBtn.graphics.newBrush(colLinesBtn.graphics.BrushType.SOLID_COLOR, [setLinesRed,setLinesGreen,setLinesBlue], 1);
        colLinesBtn.onDraw = customDraw;
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

    // Measurement text color in RGB
    var colLabel = new RGBColor;

    // Measurement lines color in RGB
    var colLines = new RGBColor;

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

        if (colLabelSet) {
            colLabelInputRed = setLabelColor(resultLabelColor)[0];
            colLabelInputGreen = setLabelColor(resultLabelColor)[1];
            colLabelInputBlue = setLabelColor(resultLabelColor)[2];
        }
        var validLabelRedColor = /^[0-9]{1,3}$/.test(colLabelInputRed) && colLabelInputRed >= 0 && colLabelInputRed <= 255;
        var validLabelGreenColor = /^[0-9]{1,3}$/.test(colLabelInputGreen) && colLabelInputGreen >= 0 && colLabelInputRed <= 255;
        var validLabelBlueColor = /^[0-9]{1,3}$/.test(colLabelInputBlue) && colLabelInputBlue >= 0 && colLabelInputBlue <= 255;
        // If colors are valid, set variables
        if (validLabelRedColor && validLabelGreenColor && validLabelBlueColor) {
            colLabel.red = colLabelInputRed;
            colLabel.green = colLabelInputGreen;
            colLabel.blue = colLabelInputBlue;
            // Set environmental variables
            $.setenv("Specify_defaultLabelColorRed", colLabel.red);
            $.setenv("Specify_defaultLabelColorGreen", colLabel.green);
            $.setenv("Specify_defaultLabelColorBlue", colLabel.blue);
            colLabelSet = false;
        }

        if (colLinesSet) {
            colLinesInputRed = setLinesColor(resultLinesColor)[0];
            colLinesInputGreen = setLinesColor(resultLinesColor)[1];
            colLinesInputBlue = setLinesColor(resultLinesColor)[2];
        }
        var validLinesRedColor = /^[0-9]{1,3}$/.test(colLinesInputRed) && colLinesInputRed >= 0 && colLinesInputRed <= 255;
        var validLinesGreenColor = /^[0-9]{1,3}$/.test(colLinesInputGreen) && colLinesInputGreen >= 0 && colLinesInputRed <= 255;
        var validLinesBlueColor = /^[0-9]{1,3}$/.test(colLinesInputBlue) && colLinesInputBlue >= 0 && colLinesInputBlue <= 255;
        // If colors are valid, set variables
        if (validLinesRedColor && validLinesGreenColor && validLinesBlueColor) {
            colLines.red = colLinesInputRed;
            colLines.green = colLinesInputGreen;
            colLines.blue = colLinesInputBlue;
            // Set environmental variables
            $.setenv("Specify_defaultLinesColorRed", colLines.red);
            $.setenv("Specify_defaultLinesColorGreen", colLines.green);
            $.setenv("Specify_defaultLinesColorBlue", colLines.blue);
            colLinesSet = false;
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
        } else if (!validLabelRedColor || !validLabelGreenColor || !validLabelBlueColor) {
            // If RGB inputs are not numeric
            beep();
            alert("Please enter a valid RGB color for the Label.");
            // colLabelInputRed.active = true;
            // colLabelInputRed.text = setLabelRed;
            // colLabelInputGreen.text = setLabelGreen;
            // colLabelInputBlue.text = setLabelBlue;
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
                var t = specLabel(w, (a + b) / 2, lines[0][1][1], colLabel);
                t.top += t.height;
            } else {
                var t = specLabel(w, (a + b) / 2, lines[0][0][1], colLabel);
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
                var t = specLabel(h, lines[0][1][0], (a + b) / 2, colLabel);
                t.left -= t.width;
                t.rotate(90, true, false, false, false, Transformation.BOTTOMRIGHT);
                t.top += t.width;
                t.top += t.height / 2;
            } else {
                var t = specLabel(h, lines[0][1][0], (a + b) / 2, colLabel);
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
            setLineStyle(p, colLines, strWidth);
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
    function specLabel(val, x, y, colLabel) {

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
        t.textRange.characterAttributes.fillColor = colLabel;

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

    function setLineStyle(path, colLines, strWidth) {
        path.filled = false;
        path.stroked = true;
        path.strokeColor = colLines;
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