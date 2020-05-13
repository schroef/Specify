/**
 * Specify
 * =================================
 * Version: 2.0.0
 * https://github.com/adamdehaven/Specify
 *
 * Adam DeHaven
 * https://adamdehaven.com
 * @adamdehaven
 *
 * Additional info:
 * https://adamdehaven.com/blog/dimension-adobe-illustrator-designs-with-a-simple-script/
 * ====================
 */

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

    // =========================================================================================== //
    // Create Dialog
    // =========================================================================================== //

    // SPECIFYDIALOGBOX
    var specifyDialogBox = new Window("dialog");
    specifyDialogBox.text = "Specify";
    specifyDialogBox.orientation = "row";
    specifyDialogBox.alignChildren = ["left", "top"];
    specifyDialogBox.spacing = 10;
    specifyDialogBox.margins = 16;

    // DIALOGMAINGROUP
    // ===============
    var dialogMainGroup = specifyDialogBox.add("group", undefined, { name: "dialogMainGroup" });
    dialogMainGroup.orientation = "column";
    dialogMainGroup.alignChildren = ["left", "center"];
    dialogMainGroup.spacing = 10;
    dialogMainGroup.margins = 0;

    // VERTICALTABBEDPANEL
    // ===================
    var verticalTabbedPanel = dialogMainGroup.add("group", undefined, undefined, { name: "verticalTabbedPanel" });
    verticalTabbedPanel.alignChildren = ["left", "fill"];
    var verticalTabbedPanel_nav = verticalTabbedPanel.add("listbox", undefined, ['OPTIONS', 'STYLES', 'RESTORE DEFAULTS']);
    var verticalTabbedPanel_innerwrap = verticalTabbedPanel.add("group")
    verticalTabbedPanel_innerwrap.alignment = ["fill", "fill"];
    verticalTabbedPanel_innerwrap.orientation = ["stack"];
    verticalTabbedPanel.alignment = ["fill", "center"];

    // TABOPTIONS
    // ==========
    var tabOptions = verticalTabbedPanel_innerwrap.add("group", undefined, { name: "tabOptions" });
    tabOptions.text = "OPTIONS";
    tabOptions.orientation = "row";
    tabOptions.alignChildren = ["left", "top"];
    tabOptions.spacing = 10;
    tabOptions.margins = 0;

    // OPTIONSMAINGROUP
    // ================
    var optionsMainGroup = tabOptions.add("group", undefined, { name: "optionsMainGroup" });
    optionsMainGroup.preferredSize.height = 410;
    optionsMainGroup.orientation = "column";
    optionsMainGroup.alignChildren = ["fill", "top"];
    optionsMainGroup.spacing = 10;
    optionsMainGroup.margins = 0;

    // DIMENSIONPANEL
    // ==============
    var dimensionPanel = optionsMainGroup.add("panel", undefined, undefined, { name: "dimensionPanel" });
    dimensionPanel.text = "Select Dimensions(s) to Specify";
    dimensionPanel.preferredSize.height = 175;
    dimensionPanel.orientation = "column";
    dimensionPanel.alignChildren = ["left", "top"];
    dimensionPanel.spacing = 10;
    dimensionPanel.margins = 20;

    var topCheckbox = dimensionPanel.add("checkbox", undefined, undefined, { name: "topCheckbox" });
    topCheckbox.helpTip = "Dimension the top side of the object(s).";
    topCheckbox.text = "Top";
    topCheckbox.alignment = ["center", "top"];
    topCheckbox.value = false;

    // DIMENSIONGROUP
    // ==============
    var dimensionGroup = dimensionPanel.add("group", undefined, { name: "dimensionGroup" });
    dimensionGroup.orientation = "row";
    dimensionGroup.alignChildren = ["center", "top"];
    dimensionGroup.spacing = 60;
    dimensionGroup.margins = 10;
    dimensionGroup.alignment = ["center", "top"];

    var leftCheckbox = dimensionGroup.add("checkbox", undefined, undefined, { name: "leftCheckbox" });
    leftCheckbox.helpTip = "Dimension the left side of the object(s).";
    leftCheckbox.text = "Left";
    leftCheckbox.value = false;

    var rightCheckbox = dimensionGroup.add("checkbox", undefined, undefined, { name: "rightCheckbox" });
    rightCheckbox.helpTip = "Dimension the right side of the object(s).";
    rightCheckbox.text = "Right";
    rightCheckbox.value = false;

    // DIMENSIONPANEL
    // ==============
    var bottomCheckbox = dimensionPanel.add("checkbox", undefined, undefined, { name: "bottomCheckbox" });
    bottomCheckbox.helpTip = "Dimension the bottom side of the object(s).";
    bottomCheckbox.text = "Bottom";
    bottomCheckbox.alignment = ["center", "top"];
    bottomCheckbox.value = false;

    var dimensionsDivider = dimensionPanel.add("panel", undefined, undefined, { name: "dimensionsDivider" });
    dimensionsDivider.alignment = "fill";

    var selectAllCheckbox = dimensionPanel.add("checkbox", undefined, undefined, { name: "selectAllCheckbox" });
    selectAllCheckbox.helpTip = "Dimension all sides of the object(s).";
    selectAllCheckbox.text = "Select All";
    selectAllCheckbox.alignment = ["center", "top"];
    selectAllCheckbox.value = false;
    selectAllCheckbox.onClick = function () {
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

    // MULTIPLEOBJECTSPANEL
    // ====================
    var multipleObjectsPanel;
    var between;

    // If exactly 2 objects are selected, give user option to dimension BETWEEN them
    if (selectedItems == 2) {
        multipleObjectsPanel = optionsMainGroup.add("panel", undefined, undefined, { name: "multipleObjectsPanel" });
        multipleObjectsPanel.text = "Multiple Objects Selected";
        multipleObjectsPanel.preferredSize.height = 65;
        multipleObjectsPanel.orientation = "column";
        multipleObjectsPanel.alignChildren = ["left", "top"];
        multipleObjectsPanel.spacing = 10;
        multipleObjectsPanel.margins = 20;

        between = multipleObjectsPanel.add("checkbox", undefined, undefined, { name: "between" });
        between.helpTip = "When checked, dimensions the distance between\nthe 2 objects for the selected dimensions.";
        between.text = "Dimension between selected objects";
        between.value = false;
    }


    // SCALEPANEL
    // ==========
    var scalePanel = optionsMainGroup.add("panel", undefined, undefined, { name: "scalePanel" });
    scalePanel.text = "Scale";
    scalePanel.preferredSize.height = 139;
    scalePanel.orientation = "column";
    scalePanel.alignChildren = ["left", "top"];
    scalePanel.spacing = 10;
    scalePanel.margins = 20;

    var customScaleInfo = scalePanel.add("statictext", undefined, undefined, { name: "customScaleInfo" });
    customScaleInfo.text = "Define the scale of the artwork/document.";

    // CUSTOMSCALEGROUP
    // ================
    var customScaleGroup = scalePanel.add("group", undefined, { name: "customScaleGroup" });
    customScaleGroup.orientation = "row";
    customScaleGroup.alignChildren = ["left", "center"];
    customScaleGroup.spacing = 10;
    customScaleGroup.margins = 0;

    var customScaleLabel = customScaleGroup.add("statictext", undefined, undefined, { name: "customScaleLabel" });
    customScaleLabel.text = "Scale:";

    var customScaleDropdown = customScaleGroup.add("dropdownlist", undefined, undefined, { name: "customScaleDropdown" });
    customScaleDropdown.helpTip = "Choose the scale of the artwork/document.\n\nExample: Choosing '1/4' will indicate the artwork is drawn at\none-fourth scale, resulting in dimension values that are 4x their\ndrawn dimensions.\n\nDefault: 1/1";
    for (var n = 1; n <= 30; n++) {
        if (n == 1) {
            customScaleDropdown.add("item", "1/" + n + "    (Default)");
            customScaleDropdown.add("separator");
        } else {
            customScaleDropdown.add("item", "1/" + n);
        }
    }
    customScaleDropdown.selection = defaultScale;
    customScaleDropdown.onChange = function () {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };

    // SCALEPANEL
    // ==========
    var scaleDivider = scalePanel.add("panel", undefined, undefined, { name: "scaleDivider" });
    scaleDivider.alignment = "fill";

    var customScaleExample = scalePanel.add("statictext", undefined, undefined, { name: "customScaleExample" });
    customScaleExample.text = "Example: 250 units at 1/4 scale displays as 1000";

    // TABSTYLES
    // =========
    var tabStyles = verticalTabbedPanel_innerwrap.add("group", undefined, { name: "tabStyles" });
    tabStyles.text = "STYLES";
    tabStyles.orientation = "column";
    tabStyles.alignChildren = ["fill", "top"];
    tabStyles.spacing = 10;
    tabStyles.margins = 0;

    // OPTIONSPANEL
    // ============
    var optionsPanel = tabStyles.add("panel", undefined, undefined, { name: "optionsPanel" });
    optionsPanel.text = "Styles";
    optionsPanel.preferredSize.height = 410;
    optionsPanel.orientation = "column";
    optionsPanel.alignChildren = ["fill", "top"];
    optionsPanel.spacing = 10;
    optionsPanel.margins = 20;

    // FONTGROUP
    // =========
    var fontGroup = optionsPanel.add("group", undefined, { name: "fontGroup" });
    fontGroup.orientation = "row";
    fontGroup.alignChildren = ["left", "center"];
    fontGroup.spacing = 10;
    fontGroup.margins = 0;

    var fontLabel = fontGroup.add("statictext", undefined, undefined, { name: "fontLabel" });
    fontLabel.text = "Font size:";

    var fontSizeInput = fontGroup.add('edittext {justify: "center", properties: {name: "fontSizeInput"}}');
    fontSizeInput.helpTip = "Enter the desired font size for the dimension label(s).\nIf value is less than one (e.g. 0.25) you must include a\nleading zero before the decimal point.\n\nDefault: " + setFontSize;
    fontSizeInput.text = defaultFontSize;
    fontSizeInput.characters = 5;
    // fontSizeInput.preferredSize.width = 35;
    fontSizeInput.onActivate = function () {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    }
    fontSizeInput.onDeactivate = function () {
        // If first character is decimal point, don't error, but instead
        // add leading zero to string.
        if (fontSizeInput.text.charAt(0) == ".") {
            fontSizeInput.text = "0" + fontSizeInput.text;
            fontSizeInput.active = true;
        }
    }

    var fontUnitsLabelText = fontGroup.add("statictext", undefined, undefined, { name: "fontUnitsLabelText" });
    fontUnitsLabelText.text = "";
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

    // DECIMALPLACESGROUP
    // ==================
    var decimalPlacesGroup = optionsPanel.add("group", undefined, { name: "decimalPlacesGroup" });
    decimalPlacesGroup.orientation = "row";
    decimalPlacesGroup.alignChildren = ["left", "center"];
    decimalPlacesGroup.spacing = 10;
    decimalPlacesGroup.margins = 0;

    var decimalPlacesLabel = decimalPlacesGroup.add("statictext", undefined, undefined, { name: "decimalPlacesLabel" });
    decimalPlacesLabel.text = "Num. of decimal places displayed:";

    var decimalPlacesInput = decimalPlacesGroup.add('edittext {justify: "center", properties: {name: "decimalPlacesInput"}}');
    decimalPlacesInput.helpTip = "Enter the desired number of decimal places to\ndisplay in the label dimensions.\n\nDefault: " + setDecimals;
    decimalPlacesInput.characters = 4;
    decimalPlacesInput.text = defaultDecimals;
    // decimalPlacesInput.preferredSize.width = 30;
    decimalPlacesInput.onActivate = function () {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
    };
    decimalPlacesInput.onChange = function () {
        decimalPlacesInput.text = decimalPlacesInput.text.replace(/[^0-9]/g, "");
    };

    // OPTIONSPANEL
    // ============
    var divider1 = optionsPanel.add("panel", undefined, undefined, { name: "divider1" });
    divider1.alignment = "fill";

    var units = optionsPanel.add("checkbox", undefined, undefined, { name: "units" });
    units.helpTip = "When checked, inserts the units label alongside\nthe outputted dimension.\nExample: 220 px";
    units.text = "Include units label(s)";
    units.value = defaultUnits;
    units.onClick = function () {
        if (units.value == false) {
            useCustomUnits.value = false;
            customUnitsInput.text = getRulerUnits();
            customUnitsInput.enabled = false;
        }
    };

    var useCustomUnits = optionsPanel.add("checkbox", undefined, undefined, { name: "useCustomUnits" });
    useCustomUnits.helpTip = "When checked, allows user to customize\nthe text of the units label.\nExample: ft";
    useCustomUnits.text = "Customize Units Label";
    useCustomUnits.value = defaultUseCustomUnits;
    useCustomUnits.onClick = function () {
        if (useCustomUnits.value == true) {
            customUnitsInput.enabled = true;
        } else {
            customUnitsInput.text = getRulerUnits();
            customUnitsInput.enabled = false;
        }
    };

    // CUSTOMUNITSGROUP
    // ================
    var customUnitsGroup = optionsPanel.add("group", undefined, { name: "customUnitsGroup" });
    customUnitsGroup.orientation = "row";
    customUnitsGroup.alignChildren = ["left", "center"];
    customUnitsGroup.spacing = 10;
    customUnitsGroup.margins = 0;

    var customUnitsLabel = customUnitsGroup.add("statictext", undefined, undefined, { name: "customUnitsLabel" });
    customUnitsLabel.text = "Custom Units Label:";

    var customUnitsInput = customUnitsGroup.add('edittext {properties: {name: "customUnitsInput", readonly: true}}');
    customUnitsInput.enabled = false;
    customUnitsInput.helpTip = "Enter the string to display after the dimension\nnumber when using a custom scale.\n\nDefault: " + setCustomUnits;
    customUnitsInput.text = defaultCustomUnits;
    // customUnitsInput.preferredSize.width = 70;
    customUnitsInput.enabled = defaultUseCustomUnits;
    customUnitsInput.characters = 20;
    customUnitsInput.onChange = function () {
        restoreDefaultsButton.enabled = true;
        infoText.enabled = true;
        customUnitsInput.text = customUnitsInput.text.replace(/[^ a-zA-Z]/g, "");
    };

    // OPTIONSPANEL
    // ============
    var divider2 = optionsPanel.add("panel", undefined, undefined, { name: "divider2" });
    divider2.alignment = "fill";

    // LABELCOLORGROUP
    // ===============
    var labelColorGroup = optionsPanel.add("group", undefined, { name: "labelColorGroup" });
    labelColorGroup.orientation = "row";
    labelColorGroup.alignChildren = ["left", "center"];
    labelColorGroup.spacing = 10;
    labelColorGroup.margins = 0;

    var colorLabel = labelColorGroup.add("statictext", undefined, undefined, { name: "colorLabel" });
    colorLabel.text = "Color:";

    var colorPicker = labelColorGroup.add("statictext", undefined, undefined, { name: "colorPicker" });
    colorPicker.text = "TODO";

    // TABDEFAULTS
    // ===========
    var tabDefaults = verticalTabbedPanel_innerwrap.add("group", undefined, { name: "tabDefaults" });
    tabDefaults.text = "RESTORE DEFAULTS";
    tabDefaults.orientation = "column";
    tabDefaults.alignChildren = ["fill", "top"];
    tabDefaults.spacing = 10;
    tabDefaults.margins = 0;

    // VERTICALTABBEDPANEL
    // ===================
    verticalTabbedPanel_tabs = [tabOptions, tabStyles, tabDefaults];

    for (var i = 0; i < verticalTabbedPanel_tabs.length; i++) {
        verticalTabbedPanel_tabs[i].alignment = ["fill", "fill"];
        verticalTabbedPanel_tabs[i].visible = false;
    }

    verticalTabbedPanel_nav.onChange = showTab_verticalTabbedPanel;

    function showTab_verticalTabbedPanel() {
        if (verticalTabbedPanel_nav.selection !== null) {
            for (var i = verticalTabbedPanel_tabs.length - 1; i >= 0; i--) {
                verticalTabbedPanel_tabs[i].visible = false;
            }
            verticalTabbedPanel_tabs[verticalTabbedPanel_nav.selection.index].visible = true;
        }
    }

    verticalTabbedPanel_nav.selection = 0; // Activate first tab
    showTab_verticalTabbedPanel()

    // DEFAULTSPANEL
    // =============
    var defaultsPanel = tabDefaults.add("panel", undefined, undefined, { name: "defaultsPanel" });
    defaultsPanel.text = "Restore Defaults";
    defaultsPanel.preferredSize.height = 410;
    defaultsPanel.orientation = "column";
    defaultsPanel.alignChildren = ["center", "center"];
    defaultsPanel.spacing = 10;
    defaultsPanel.margins = 10;

    // RESTOREDEFAULTSGROUP
    // ====================
    var restoreDefaultsGroup = defaultsPanel.add("group", undefined, { name: "restoreDefaultsGroup" });
    restoreDefaultsGroup.orientation = "column";
    restoreDefaultsGroup.alignChildren = ["left", "center"];
    restoreDefaultsGroup.spacing = 10;
    restoreDefaultsGroup.margins = 0;

    var infoText = restoreDefaultsGroup.add("statictext", undefined, undefined, { name: "infoText" });
    infoText.text = "Options are persistent until application is closed.";

    var restoreDefaultsButton = restoreDefaultsGroup.add("button", undefined, undefined, { name: "restoreDefaultsButton" });
    restoreDefaultsButton.text = "Restore All Defaults";
    restoreDefaultsButton.preferredSize.width = 124;
    restoreDefaultsButton.alignment = ["center", "center"];
    restoreDefaultsButton.enabled = (setFontSize != defaultFontSize || setRed != defaultColorRed || setGreen != defaultColorGreen || setBlue != defaultColorBlue || setDecimals != defaultDecimals || setScale != defaultScale || setCustomUnits != defaultCustomUnits ? true : false);
    restoreDefaultsButton.onClick = function () {
        restoreDefaults();
    };

    function restoreDefaults() {
        units.value = setUnits;
        fontSizeInput.text = setFontSize;
        colorInputRed.text = setRed;
        colorInputGreen.text = setGreen;
        colorInputBlue.text = setBlue;
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
        $.setenv("Specify_defaultDecimals", "");
        $.setenv("Specify_defaultScale", "");
        $.setenv("Specify_defaultUseCustomUnits", "");
        $.setenv("Specify_defaultCustomUnits", "");
    };

    // FOOTERGROUP
    // ===========
    var footerGroup = dialogMainGroup.add("group", undefined, { name: "footerGroup" });
    footerGroup.orientation = "row";
    footerGroup.alignChildren = ["left", "center"];
    footerGroup.spacing = 20;
    footerGroup.margins = 0;

    // UPDATESGROUP
    // ============
    var updatesGroup = footerGroup.add("group", undefined, { name: "updatesGroup" });
    updatesGroup.orientation = "column";
    updatesGroup.alignChildren = ["left", "center"];
    updatesGroup.spacing = 3;
    updatesGroup.margins = 0;

    var specifyUpdatesText = updatesGroup.add("statictext", undefined, undefined, { name: "specifyUpdatesText" });
    specifyUpdatesText.text = "For updates & more info:";

    var urlInput = updatesGroup.add('edittext {justify: "center", properties: {name: "urlInput", readonly: true}}');
    urlInput.text = "https://github.com/adamdehaven/Specify";

    // BUTTONGROUP
    // ===========
    var buttonGroup = footerGroup.add("group", undefined, { name: "buttonGroup" });
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = ["right", "center"];
    buttonGroup.spacing = 10;
    buttonGroup.margins = 0;

    var cancelButton = buttonGroup.add("button", undefined, undefined, { name: "cancelButton" });
    cancelButton.text = "Cancel";
    cancelButton.onClick = function () {
        specifyDialogBox.close();
    };

    var specifyButton = buttonGroup.add("button", undefined, undefined, { name: "specifyButton" });
    specifyButton.text = "Specify Object(s)";
    specifyButton.preferredSize.height = 40;
    // specifyDialogBox.defaultElement = specifyButton;
    specifyButton.onClick = function () {
        startSpec();
    };

    // =========================================================================================== //
    // END: Create Dialog
    // =========================================================================================== //

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

    // Gap between measurement lines and object
    var gap = 4;

    // Size of perpendicular measurement lines.
    var size = 6;

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
        // // If colors are valid, set variables
        // if (validRedColor && validGreenColor && validBlueColor) {
        //     color.red = colorInputRed.text;
        //     color.green = colorInputGreen.text;
        //     color.blue = colorInputBlue.text;
        //     // Set environmental variables
        //     $.setenv("Specify_defaultColorRed", color.red);
        //     $.setenv("Specify_defaultColorGreen", color.green);
        //     $.setenv("Specify_defaultColorBlue", color.blue);
        // }

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
        } /*else if (!validRedColor || !validGreenColor || !validBlueColor) {
            // If RGB inputs are not numeric
            beep();
            alert("Please enter a valid RGB color.");
            colorInputRed.active = true;
            colorInputRed.text = defaultColorRed;
            colorInputGreen.text = defaultColorGreen;
            colorInputBlue.text = defaultColorBlue;
        }*/ else if (!validDecimalPlaces) {
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
            lines[0] = new Array(new Array(a, c + (gap) * dir));
            lines[0].push(new Array(a, c + (gap + size) * dir));
            lines[1] = new Array(new Array(b, c + (gap) * dir));
            lines[1].push(new Array(b, c + (gap + size) * dir));

            // 1 horizontal line
            lines[2] = new Array(new Array(a, c + (gap + size / 2) * dir));
            lines[2].push(new Array(b, c + (gap + size / 2) * dir));

            // Create text label
            if (where == "Top") {
                var t = specLabel(w, (a + b) / 2, lines[0][1][1], color);
                t.top += t.height;
            } else {
                var t = specLabel(w, (a + b) / 2, lines[0][0][1], color);
                t.top -= size;
            }
            t.left -= t.width / 2;

        } else {
            // Vertical measurement

            // 2 horizontal lines
            lines[0] = new Array(new Array(c + (gap) * dir, a));
            lines[0].push(new Array(c + (gap + size) * dir, a));
            lines[1] = new Array(new Array(c + (gap) * dir, b));
            lines[1].push(new Array(c + (gap + size) * dir, b));

            //1 vertical line
            lines[2] = new Array(new Array(c + (gap + size / 2) * dir, a));
            lines[2].push(new Array(c + (gap + size / 2) * dir, b));

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
            setLineStyle(p, color);
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
        var labelFontInUnits = convertToPoints(labelFontSize);

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

    function setLineStyle(path, color) {
        path.filled = false;
        path.stroked = true;
        path.strokeColor = color;
        path.strokeWidth = 0.5;
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
