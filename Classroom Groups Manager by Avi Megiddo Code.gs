const badPairs = new Set([]);
const forbiddenWildCards = new Set([]);
function onOpen() {
 var ui = SlidesApp.getUi();
 // Add a custom menu to Google Slides.
 ui.createMenu('Classroom Groups Manager')
   .addItem('Open Sidebar', 'showSidebar')
   .addToUi();
}

function showSidebar() {
 var html = HtmlService.createHtmlOutputFromFile('Sidebar')
   .setTitle('Classroom Groups Manager')  // Update this to your new title
   .setWidth(200);
 SlidesApp.getUi().showSidebar(html);
}


var RECT_WIDTH = 0.9 * 100;  // in points
var RECT_HEIGHT = 0.66 * 100;  // in points
var GAP_BETWEEN_GROUPS = 5;  // in points, adjust as needed


const scriptProperties = PropertiesService.getScriptProperties();


function createNamedRectangles(names) {
 var ui = SlidesApp.getUi(); // For user alerts and debug
 Logger.log("Function createNamedRectangles called.");

 // Clean up names: remove empty or invalid entries (such as empty strings or spaces)
 names = names.filter(function(name) {
   return name && name.trim() !== "";
 });

 // Check if valid names are provided after filtering
 if (!names || names.length === 0) {
   ui.alert("Error: No valid names provided. Please enter valid student names before proceeding.");
   return;
 }

 var presentation = SlidesApp.getActivePresentation();
 var slide = presentation.getSelection().getCurrentPage().asSlide();

 // Clear all shapes from the slide
 slide.getShapes().forEach(function (shape) {
   shape.remove();
 });

 var xPosition = 20;
 var yPosition = 50;
 var columnCounter = 0;

 for (var i = 0; i < names.length; i++) {
   try {
     // Insert a round rectangle for each valid name
     var shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPosition, yPosition, RECT_WIDTH, RECT_HEIGHT);

     // Set a random pastel color
     var randomColor = getRandomPastelColor();
     shape.getFill().setSolidFill(randomColor);

     var textRange = shape.getText();
     textRange.setText(names[i]);
     textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

     var textStyle = textRange.getTextStyle();
     textStyle.setFontFamily('Barlow');

     // Adjust font size based on the length of the name
     var fontSize = 16;
     if (names[i].length >= 7) fontSize = 14;
     if (names[i].length >= 10) fontSize = 12;

     textStyle.setFontSize(fontSize);
     textStyle.setForegroundColor('#000000');

     // Ensure text padding (insets) is zero
     var paragraphStyle = textRange.getParagraphStyle();
     paragraphStyle.setSpaceAbove(0).setSpaceBelow(0);
     shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

   } catch (e) {
     // Provide meaningful error messages in the UI
     ui.alert(`Error inserting shape for name at position ${i + 1}.\nName: "${names[i]}"\nError: ${e.message}`);
     Logger.log(`Error inserting shape for name: ${names[i]}, Error: ${e.message}`);
     return;  // Exit the function on error
   }

   // Update positioning for the next shape
   yPosition += RECT_HEIGHT + 20;
   columnCounter++;

   // Shift to a new column after every 5 shapes
   if (columnCounter === 5) {
     yPosition = 50;
     xPosition += RECT_WIDTH + GAP_BETWEEN_GROUPS + 20;
     columnCounter = 0;
   }
 }

 ui.alert("All desks created successfully!");

 // After creating all the rectangles, save the names to Script Properties
 var slideId = slide.getObjectId();
 scriptProperties.setProperty(slideId, JSON.stringify(names));

 // Open the dialog box to label the class
 openDialog();
}







function openDialog() {
 var html = HtmlService.createHtmlOutputFromFile('showClassNameDialog')
   .setWidth(500)
   .setHeight(400);
 SlidesApp.getUi().showModalDialog(html, 'Enter Class Name/Label');

}


function addClassLabelToSlide(className) {
 try {
   if (!className) {
     return;  // Exit if no class name provided
   }

   var presentation = SlidesApp.getActivePresentation();
   var slide = presentation.getSelection().getCurrentPage().asSlide();

   // Dimensions and position of the label
   var labelWidth = 300;
   var labelHeight = 50;
   var xPosition = presentation.getPageWidth() - labelWidth - 10;
   var yPosition = presentation.getPageHeight() - labelHeight - 10;

   var shape = slide.insertShape(SlidesApp.ShapeType.PLAQUE, xPosition, yPosition, labelWidth, labelHeight);
   var shapeId = shape.getObjectId();
   slide.getNotesPage().getSpeakerNotesShape().getText().setText("CLASS_LABEL_ID: " + shapeId);

   // Send shape to back layer
   shape.sendToBack();

   // Set text within the shape
   var textRange = shape.getText();
   textRange.setText(className);
   textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

   // Style the text
   var textStyle = textRange.getTextStyle();
   textStyle.setFontFamily('Barlow');
   textStyle.setForegroundColor('#000000');

   // Function to resize the text if it's too big for the shape
   function resizeTextToFit() {
     var fontSize = 36;  // Start with a larger font size
     var maxFontSize = 36;  // Maximum font size
     var minFontSize = 10;  // Minimum font size

     // Resize the text to fit within the shape's width
     while (fontSize >= minFontSize) {
       textStyle.setFontSize(fontSize);

       // Check if the text fits within the shape by comparing width
       var shapeTextWidth = shape.getText().getTextStyle().getFontSize();
       if (shape.getText().asString().length * fontSize / 2 < labelWidth - 10) {  // Approximate text width
         break;  // Text fits, exit loop
       }
       fontSize -= 1;  // Reduce font size and try again
     }

     // Set final font size
     textStyle.setFontSize(fontSize);
   }

   // Call the function to resize the text if needed
   resizeTextToFit();

   // Vertically align the text in the middle of the rectangle
   var paragraphStyle = textRange.getParagraphStyle();
   paragraphStyle.setSpaceAbove((labelHeight - 30) / 2);
   paragraphStyle.setSpaceBelow((labelHeight - 30) / 2);

   return "Added label: " + className;
 } catch (e) {
   return "An error occurred: " + e.toString();
 }
}



function clearTextArea() {
 var textBox = DocumentApp.getActiveDocument().getBody().findText('YOUR_TEXT_BOX_IDENTIFIER');
 if (textBox) {
   textBox.clear();
 }
}


function getStoredNamesForCurrentSlide() {
 var slideId = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId();
 var storedNames = PropertiesService.getScriptProperties().getProperty(slideId);
 if (storedNames) {
   return JSON.parse(storedNames);
 } else {
   return null;
 }
}

function resetWholeClass() {

 var ui = SlidesApp.getUi(); // For dialog boxes
 var presentation = SlidesApp.getActivePresentation();
 var slide = presentation.getSelection().getCurrentPage().asSlide();
 var slideId = slide.getObjectId();  // Get the ID of the current slide
 var names = [];

 // Step 1: Exclude the class label shape using the CLASS_LABEL_ID from speaker notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 // Step 2: Check for names on the current slide's rectangles, excluding the class label
 var shapes = slide.getShapes().filter(function (shape) {
   return shape.getObjectId() !== classLabelId;  // Exclude class label shape
 });

 names = shapes.map(function (shape) {
   return shape.getText().asString().trim();
 }).filter(Boolean);  // Remove empty entries

 if (names.length > 0) {
   ui.alert('Found names on the slide. Resetting desks based on these names.');
   createNamedRectangles(names);
   return names.join(", ");
 }

 // Step 3: If no names on the slide, check script properties (via slide ID)
 var namesString = scriptProperties.getProperty(slideId);

 if (namesString) {
   names = JSON.parse(namesString);
   if (names.length > 0) {
     ui.alert('Retrieved saved names from script properties.');
     createNamedRectangles(names);
     return names.join(", ");
   }
 }

 // Step 4: If no names in script properties, check class ID in speaker notes for stored names
 if (classLabelId) {
   var classLabelNamesString = scriptProperties.getProperty(classLabelId);
   if (classLabelNamesString) {
     names = JSON.parse(classLabelNamesString);
     if (names.length > 0) {
       ui.alert('Found names stored under class label ID from speaker notes.');
       createNamedRectangles(names);
       return names.join(", ");
     }
   }
 }

 // Step 5: If all else fails, check the sidebar (you would need to implement logic to retrieve from sidebar)
 var sidebarNames = getNamesFromSidebar();  // Placeholder function for getting sidebar input
 if (sidebarNames && sidebarNames.length > 0) {
   ui.alert('Found names in the sidebar.');
   createNamedRectangles(sidebarNames);
   return sidebarNames.join(", ");
 }

 // If no names are found from any source, display an error message
 ui.alert("Sorry, no data found for this class.");
 return "";
}

// Placeholder function for retrieving names from the sidebar
function getNamesFromSidebar() {
 // Implement the logic to retrieve the names from the sidebar here
 // For example, this could read input from an HTML sidebar element and return it as an array
 return [];  // For now, returning an empty array
}


// Placeholder function for retrieving names from the sidebar
function getNamesFromSidebar() {
 // Implement the logic to retrieve the names from the sidebar here
 // For example, this could read input from an HTML sidebar element and return it as an array
 return [];  // For now, returning an empty array
}


function assignNamesToDesks() {
 var ui = SlidesApp.getUi(); // For alerts

 // Get the list of names from the script properties

 var studentNamesJson = scriptProperties.getProperty("studentNames");
 var studentNames = JSON.parse(studentNamesJson) || [];

 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 clearWildCardStyling(shapes);

 // Read the class label ID from the speaker notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 // Filter out the class label and any shapes without text
 var filteredShapes = shapes.filter(function (shape) {
   var shapeId = shape.getObjectId();
   return shapeId !== classLabelId;
 });

 // Extract names from shapes early
 var shapeNames = filteredShapes.map(function (shape) {
   return shape.getText().asString().trim(); // Get names from shapes
 }).filter(Boolean); // Filter out any empty or null entries

 // If student names from shapes are found, use them
 if (shapeNames.length > 0) {
   studentNames = shapeNames;
 }

 // If no student names are found from shapes or script properties, alert the user
 if (studentNames.length === 0) {
   ui.alert("No names found in shapes or script properties.");
   return;
 }

 // Shuffle the student names
 shuffleArray(studentNames);

 // Clear existing names in the shapes
 filteredShapes.forEach(function (shape) {
   shape.getText().setText(''); // Clear previous text
 });

 // Assign new names to shapes with text resizing logic
 for (var i = 0; i < Math.min(filteredShapes.length, studentNames.length); i++) {
   var shape = filteredShapes[i];

   var textRange = shape.getText();
   textRange.setText(studentNames[i]); // Assign the name to the shape
   textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

   var textStyle = textRange.getTextStyle();
   textStyle.setFontFamily('Barlow');

   // Apply text resizing logic based on the length of the student's name (greater than or equal to)
   var fontSize = 14;  // Start with 14pt font
   if (studentNames[i].length >= 7) {
     fontSize = 11;
   }
   if (studentNames[i].length >= 10) {
     fontSize = 10;
   }

   textStyle.setFontSize(fontSize);
   textStyle.setForegroundColor('#000000');
 }

 ui.alert("Names assigned to desks successfully!");
}






function setStudentNames(namesArray) {
 var scriptProperties = PropertiesService.getScriptProperties();
 scriptProperties.setProperty("studentNames", JSON.stringify(namesArray));
 var ui = SpreadsheetApp.getUi();
 ui.alert('Debug', 'Student Names set in Properties: ' + JSON.stringify(namesArray), ui.ButtonSet.OK);  // Debug step 1
}




function groupInPairs() {
 scriptProperties.setProperty('resizedShapes', JSON.stringify({}));

 loadBadPairs();
 loadForbiddenWildCards();

 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 clearWildCardStyling(shapes);

 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 shapes.forEach(function (shape) {

   shape.setWidth(RECT_WIDTH);
   shape.setHeight(RECT_HEIGHT);


   var textRange = shape.getText();
   textRange.getTextStyle().setBold(false);
   shape.getFill().setSolidFill("#FFFFFF");
 });

 var containsBadPair, containsBadWildcard;

 do {
   // Loop Logic
   shuffleArray(shapes);
   containsBadPair = false;
   containsBadWildcard = false;

   // Validation Methods
   for (var i = 0; i < shapes.length; i += 2) {
     if (shapes[i] && shapes[i + 1]) {
       if (isBadPair(shapes[i], shapes[i + 1])) {
         containsBadPair = true;
         break;
       }
     }
   }

   if (!containsBadPair) {
     var remainderShapes = [];
     var remainder = shapes.length % 2;
     for (var j = shapes.length - remainder; j < shapes.length; j++) {
       remainderShapes.push(shapes[j]);
     }

     for (const shape of remainderShapes) {
       const name = shape.getText().asString().trim().replace(/\n/g, '');
       if (forbiddenWildCards.has(name)) {
         console.log(`Found forbidden wildcard: ${name}`);
         containsBadWildcard = true;
         break;
       }
     }
   }


 } while (containsBadPair || containsBadWildcard);

 // 3. Correct `shapes` Array
 // Existing logic for using the `shapes` array

 // 4. Function Calls
 initializeAvailableColors();

 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 for (var i = 0; i < shapes.length; i += 2) {
   var color = getRandomPastelColor();  // Get a random color for the group

   if (shapes[i]) {
     shapes[i].setTop(yPosition).setLeft(xPosition);
     shapes[i].getFill().setSolidFill(color);
   }
   if (shapes[i + 1]) {
     shapes[i + 1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH);
     shapes[i + 1].getFill().setSolidFill(color);
   }

   groupCounter++;
   if (groupCounter % 5 == 0) {
     yPosition = 20;
     xPosition += 2.7 * RECT_WIDTH + GAP_BETWEEN_GROUPS;
   } else {
     yPosition += RECT_HEIGHT + GAP_BETWEEN_GROUPS + 20;
   }
 }
 // Create an array to hold the remainder shapes
 var remainderShapes = [];

 // Calculate the remainder for groupInFours
 var remainder = shapes.length % 2;

 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 // Loop to the length minus remainder to place shapes normally
 for (var i = 0; i < shapes.length - remainder; i += 4) {
   // ... (existing code to place and color shapes) ...
 }

 // Now collect the remainder shapes in the remainderShapes array
 for (var j = shapes.length - remainder; j < shapes.length; j++) {
   remainderShapes.push(shapes[j]);
 }

 // Call labelWildCard to style the remainder shapes
 if (remainder > 0) {
   labelWildCard(remainderShapes);
 }
}


function groupInThrees() {
 scriptProperties.setProperty('resizedShapes', JSON.stringify({}));
 loadBadPairs();
 loadForbiddenWildCards();
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 clearWildCardStyling(shapes);

 // Get the class label shape ID from the notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 // Filter out the class label shape
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 // Reset the font and fill for each rectangle
 shapes.forEach(function (shape) {
   shape.setWidth(RECT_WIDTH);
   shape.setHeight(RECT_HEIGHT);
   var textRange = shape.getText();
   textRange.getTextStyle().setBold(false);
   shape.getFill().setSolidFill("#FFFFFF");  // Reset to white or your default color
 });


 // Now, filter out the shapes whose ID matches the class label shape's ID
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 var containsBadGroup;
 var containsBadWildcard;

 do {
   shuffleArray(shapes);
   containsBadGroup = false;

   for (let i = 0; i < shapes.length; i += 3) {
     let group = shapes.slice(i, i + 3);
     if (containsBadPairInGroup(group)) {
       containsBadGroup = true;
       break;
     }
   }


   // Validation for forbidden wildcards
   if (!containsBadGroup) {
     var remainderShapes = [];
     var remainder = shapes.length % 3;
     for (let j = shapes.length - remainder; j < shapes.length; j++) {
       remainderShapes.push(shapes[j]);
     }

     for (const shape of remainderShapes) {
       const name = shape.getText().asString().trim().replace(/\n/g, '');
       if (forbiddenWildCards.has(name)) {
         console.log(`Found forbidden wildcard: ${name}`);
         containsBadWildcard = true;
         break;
       }
     }
   }
 } while (containsBadGroup || containsBadWildcard);

 initializeAvailableColors();  // Initialize the colors at the start of function

 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 for (var i = 0; i < shapes.length; i += 3) {
   var color = getRandomPastelColor();  // Get a random color for the group

   if (shapes[i]) {
     shapes[i].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
   }
   if (shapes[i + 1]) {
     shapes[i + 1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
   }
   if (shapes[i + 2]) {
     shapes[i + 2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + 0.5 * RECT_WIDTH).getFill().setSolidFill(color);
   }

   groupCounter++;
   if (groupCounter % 3 == 0) {
     yPosition = 20;
     xPosition += 2.7 * RECT_WIDTH + GAP_BETWEEN_GROUPS;
   } else {
     yPosition += 2.3 * RECT_HEIGHT + GAP_BETWEEN_GROUPS;
   }
 }

 // Create an array to hold the remainder shapes
 var remainderShapes = [];

 // Calculate the remainder for groupInFours
 var remainder = shapes.length % 3;

 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 // Loop to the length minus remainder to place shapes normally
 for (var i = 0; i < shapes.length - remainder; i += 4) {
   // ... (existing code to place and color shapes) ...
 }

 // Now collect the remainder shapes in the remainderShapes array
 for (var j = shapes.length - remainder; j < shapes.length; j++) {
   remainderShapes.push(shapes[j]);
 }

 // Call labelWildCard to style the remainder shapes
 if (remainder > 0) {
   labelWildCard(remainderShapes);
 }
}


function groupInFours() {
 scriptProperties.setProperty('resizedShapes', JSON.stringify({}));
 loadBadPairs();
 loadForbiddenWildCards();
 deselectStudent(); // Deselect previously selected student if any
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 clearWildCardStyling(shapes);


 // Get the class label shape ID from the notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 // Filter out the class label shape
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 // Reset the font and fill for each rectangle
 shapes.forEach(function (shape) {
   shape.setWidth(RECT_WIDTH);
   shape.setHeight(RECT_HEIGHT);
   var textRange = shape.getText();
   textRange.getTextStyle().setBold(false);
   shape.getFill().setSolidFill("#FFFFFF");  // Reset to white or your default color
 });


 // Now, filter out the shapes whose ID matches the class label shape's ID
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });


 var containsBadGroup;
 var containsBadWildcard;

 do {
   shuffleArray(shapes);
   containsBadGroup = false;

   for (let i = 0; i < shapes.length; i += 4) {
     let group = shapes.slice(i, i + 4);
     if (containsBadPairInGroup(group)) {
       containsBadGroup = true;
       break;
     }
   }


   // Validation for forbidden wildcards
   if (!containsBadGroup) {
     var remainderShapes = [];
     var remainder = shapes.length % 4;
     for (let j = shapes.length - remainder; j < shapes.length; j++) {
       remainderShapes.push(shapes[j]);
     }

     for (const shape of remainderShapes) {
       const name = shape.getText().asString().trim().replace(/\n/g, '');
       if (forbiddenWildCards.has(name)) {
         console.log(`Found forbidden wildcard: ${name}`);
         containsBadWildcard = true;
         break;
       }
     }
   }
 } while (containsBadGroup || containsBadWildcard);

 initializeAvailableColors();  // Initialize the colors at the start of function

 var xPosition = 50;
 var yPosition = 50;
 var groupCounter = 0;

 for (var i = 0; i < shapes.length; i += 4) {
   var color = getRandomPastelColor();  // Get a random color for the group

   if (shapes[i]) {
     shapes[i].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
   }
   if (shapes[i + 1]) {
     shapes[i + 1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
   }
   if (shapes[i + 2]) {
     shapes[i + 2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition).getFill().setSolidFill(color);
   }
   if (shapes[i + 3]) {
     shapes[i + 3].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
   }

   groupCounter++;
   if (groupCounter % 3 == 0) {
     yPosition = 50;
     xPosition += 2 * RECT_WIDTH + GAP_BETWEEN_GROUPS + 20;
   } else {
     yPosition += 2 * RECT_HEIGHT + GAP_BETWEEN_GROUPS + 20;
   }
 }
 // Create an array to hold the remainder shapes
 var remainderShapes = [];

 // Calculate the remainder for groupInFours
 var remainder = shapes.length % 4;

 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 // Loop to the length minus remainder to place shapes normally
 for (var i = 0; i < shapes.length - remainder; i += 4) {
   // ... (existing code to place and color shapes) ...
 }

 // Now collect the remainder shapes in the remainderShapes array
 for (var j = shapes.length - remainder; j < shapes.length; j++) {
   remainderShapes.push(shapes[j]);
 }

 // Call labelWildCard to style the remainder shapes
 if (remainder > 0) {
   labelWildCard(remainderShapes);
 }
}


function labelWildCard(remainderShapes) {
 remainderShapes.forEach(function (shape) {
   shape.getFill().setSolidFill("#FFD700"); // Gold background for wild card
   try {
     shape.getBorder().setWeight(4).getLineFill().setSolidFill('#F0E68C'); // Light gold border (thick) only for wild card
   } catch (e) {
     Logger.log('Shape does not support border styling: ' + e.message); // Log the error for debugging
   }
   var textRange = shape.getText();
   if (textRange) {
     textRange.getTextStyle().setBold(true); // Bold text for wild card
   }
 });
}



function clearWildCardStyling(shapes) {
 shapes.forEach(function (shape) {
   try {
     // Reset fill color to default (white or whatever base color you use)
     shape.getFill().setSolidFill("#FFFFFF");

     // Ensure text style is reset, especially removing bold
     var textRange = shape.getText();
     if (textRange) {
       textRange.getTextStyle().setBold(false);  // Reset bold style
     }

     // Reset the border to thinner (e.g., 0.5pt thickness)
     if (shape.getBorder) {
       var border = shape.getBorder();
       if (border) {
         border.setWeight(0.5);  // Set border thickness to 0.5pt
         border.getLineFill().setSolidFill("#000000");  // Set border color to black
       } else {
         Logger.log("Shape does not have a border to reset.");
       }
     } else {
       Logger.log("Shape does not support border.");
     }

     Logger.log("Shape " + shape.getObjectId() + " border and styling reset successfully.");
    
   } catch (e) {
     Logger.log("Error resetting wild card styling for shape " + shape.getObjectId() + ": " + e.message);
   }
 });
}











function groupInFives() {
 scriptProperties.setProperty('resizedShapes', JSON.stringify({}));
 loadBadPairs();
 loadForbiddenWildCards();
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();
 clearWildCardStyling(shapes);

 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 var numStudents = shapes.length;

 shapes.forEach(function (shape) {
   shape.setWidth(RECT_WIDTH);
   shape.setHeight(RECT_HEIGHT);
   var textRange = shape.getText();
   textRange.getTextStyle().setBold(false);
   shape.getFill().setSolidFill("#FFFFFF");
 });

 shuffleArray(shapes);
 initializeAvailableColors();

 var xPosition = 50;
 var yPosition = 50;
 var groupCounter = 0;
 var remainder = numStudents % 5;
 var GAP_BETWEEN_GROUPS = 30; // Adjust this to increase the vertical and horizontal spacing between groups

 // Loop for placing students in groups of 5
 for (var i = 0; i < numStudents - remainder; i += 5) {
   var color = getRandomPastelColor();

   for (var j = 0; j < 5; j++) {
     var currentXPosition = xPosition + (j % 2) * RECT_WIDTH; // X position logic for first four desks
     var currentYPosition = yPosition;

     // For the fifth desk (index 4), center it between the two rows
     if (j === 4) {
       currentYPosition = yPosition + RECT_HEIGHT / 2; // Center the fifth desk vertically
       currentXPosition = xPosition + RECT_WIDTH * 2; // Place to the right of the other four
     } else if (j >= 2) {
       currentYPosition += RECT_HEIGHT; // Move to the bottom row for the third and fourth desk
     }

     shapes[i + j].setTop(currentYPosition).setLeft(currentXPosition);
     shapes[i + j].getFill().setSolidFill(color);
   }

   groupCounter++;
   if (groupCounter % 3 === 0) {
     yPosition = 50; // Reset Y position after 3 groups
     xPosition += 3 * RECT_WIDTH + GAP_BETWEEN_GROUPS; // Shift to the right for the next group
   } else {
     yPosition += 2 * RECT_HEIGHT + GAP_BETWEEN_GROUPS; // Adjust Y position for the next group with a gap
   }
 }

 // Handle the remainder shapes separately at the end
 if (remainder > 0) {
   var remainderShapes = [];
   var color = getRandomPastelColor();
   var remainderX = xPosition;
   var remainderY = yPosition;

   for (var i = numStudents - remainder; i < numStudents; i++) {
     remainderShapes.push(shapes[i]);

     // Adjust placement to avoid overlap
     var localIndex = i - (numStudents - remainder);
     var row = Math.floor(localIndex / 3);
     var col = localIndex % 3;

     var finalX = remainderX + col * RECT_WIDTH;
     var finalY = remainderY + row * RECT_HEIGHT;

     shapes[i].setTop(finalY).setLeft(finalX);
     shapes[i].getFill().setSolidFill(color);
   }

   labelWildCard(remainderShapes);
 }
}

function loadBadPairs() {
 const badPairsString = scriptProperties.getProperty('badPairs') || "";
 const pairs = badPairsString.split(',').map(pair => pair.trim());

 badPairs.clear();

 pairs.forEach(pair => {
   badPairs.add(pair);  // No need to add both combinations, as the pair is already sorted
 });
}


function loadForbiddenWildCards() {
 forbiddenWildCards.clear();

 for (const pair of badPairs) {
   const [name1, name2] = pair.split(' ');
   if (name1 && name2) {
     forbiddenWildCards.add(name1.trim());
     forbiddenWildCards.add(name2.trim());
   } else {
     console.log(`Malformed pair skipped: ${pair}`);
   }
 }
}



function saveBadPair(badPairString) {
 let existingBadPairs = scriptProperties.getProperty('badPairs');
 existingBadPairs = existingBadPairs ? existingBadPairs.split(',') : [];

 // Split and trim the pair names using spaces or commas as delimiters
 const pairs = badPairString.split(/[,\s]+/);

 // Ensure both names are defined before adding to the bad pairs list
 if (pairs.length >= 2) {
   const [name1, name2] = pairs.map(pair => pair.trim());

   // Sort the names alphabetically and create a single canonical form
   const sortedPair = [name1, name2].sort().join(' ');

   // Check if the pair already exists to avoid duplication
   if (!existingBadPairs.includes(sortedPair)) {
     existingBadPairs.push(sortedPair);
   }

   // Save the updated list of bad pairs
   scriptProperties.setProperty('badPairs', existingBadPairs.join(','));

   // Display updated list in a Slides modal dialog
   const ui = SlidesApp.getUi();
   ui.alert('Updated bad pairs list:\n' + existingBadPairs.join('\n'));
 }
}


function saveBadPairs(badPairsArray) {


 // Convert the array of bad pairs back into a comma-separated string
 const badPairsString = badPairsArray.join(',');

 Logger.log("Saving updated bad pairs: " + badPairsString);
 scriptProperties.setProperty('badPairs', badPairsString);  // Save the string format

 // Optionally show a toast or modal confirming the update
 SlidesApp.getUi().alert('Bad pairs updated successfully!');
}


function getBadPairs() {

 const badPairsString = scriptProperties.getProperty('badPairs') || '';

 // Split the string by commas to get the individual pairs
 const badPairsArray = badPairsString.split(',').filter(Boolean); // Filter out empty strings
 Logger.log("Retrieved bad pairs: " + badPairsArray);
 return badPairsArray;
}


function isBadPair(shape1, shape2) {
 const badPairs = getBadPairs();  // This returns an array
 const name1 = shape1.getText().asString().trim().replace(/\n/g, '');
 const name2 = shape2.getText().asString().trim().replace(/\n/g, '');

 // Create the sorted pair for comparison
 const sortedPair = [name1, name2].sort().join(' ');

 // Check if the sorted pair exists in the bad pairs (array), using includes() instead of has()
 return badPairs.includes(sortedPair);
}

function containsBadPairInGroup(shapes) {
 // Loop through all possible pair combinations in the array of shapes
 for (let i = 0; i < shapes.length; i++) {
   for (let j = i + 1; j < shapes.length; j++) {
     if (isBadPair(shapes[i], shapes[j])) {
       return true;
     }
   }
 }
 return false;
}

function editBadPairings() {

 const badPairsString = scriptProperties.getProperty('badPairs') || '';

 // Split the comma-separated string into an array
 const badPairsArray = badPairsString.split(',').filter(Boolean);  // Filter out empty strings

 Logger.log("Bad pairs array: " + JSON.stringify(badPairsArray));

 // Pass the bad pairs as a properly formatted string (manually JSON.stringify it)
 const template = HtmlService.createTemplateFromFile('EditBadPairsDialog');
 template.badPairs = JSON.stringify(badPairsArray);  // Manually pass as JSON string

 const htmlOutput = template.evaluate()
   .setWidth(400)
   .setHeight(600);

 SlidesApp.getUi().showModalDialog(htmlOutput, 'Edit Bad Pairings');
}



function clearBadPairs() {

 scriptProperties.deleteProperty('badPairs');  // This will clear the bad pairs
}

function groupAroundSelectedStudents() {
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();

 // Get the selected student list from the sidebar input
 var selectedStudentsInput = document.getElementById("selectedStudents").value;
 var selectedStudents = selectedStudentsInput.split(',').map(s => s.trim());

 // Get the class label shape ID from the speaker notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 var shapes = slide.getShapes();

 // Filter out the class label shape
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 // Shuffle shapes to ensure random distribution
 shuffleArray(shapes);

 // Create a number of empty groups equal to the number of selected students
 var groups = new Array(selectedStudents.length).fill(null).map(() => []);

 // Step 1: Place each selected student in their own group
 selectedStudents.forEach(function (student, index) {
   var studentShape = shapes.find(function (shape) {
     return shape.getText().asString().trim() === student;
   });

   if (studentShape) {
     groups[index].push(studentShape);
     shapes = shapes.filter(function (shape) {
       return shape !== studentShape;
     });
   } else {
     Logger.log("Could not find shape for student: " + student);
   }
 });

 // Step 2: Distribute the remaining students randomly among the groups
 var groupIndex = 0;
 shapes.forEach(function (shape) {
   groups[groupIndex].push(shape);
   groupIndex = (groupIndex + 1) % selectedStudents.length;
 });

 // Step 3: Arrange the groups on the slide
 arrangeGroupsOnSlide(groups);
}

function groupAroundSelectedStudentsFromSidebar(selectedStudentsInput) {
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();

 // Debug: Alert the input received from the sidebar
 SlidesApp.getUi().alert('Selected Students from Sidebar: ' + selectedStudentsInput);

 // Split the input into an array of students
 var selectedStudents = selectedStudentsInput.split(/[\s,]+/).map(s => s.trim());


 // Debug: Alert the processed list of selected students
 SlidesApp.getUi().alert('Processed Selected Students: ' + selectedStudents.join(', '));

 // Get the class label shape ID from the speaker notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 var shapes = slide.getShapes();

 // Filter out the class label shape
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 // Debug: Alert the number of student shapes found on the slide
 SlidesApp.getUi().alert('Number of student shapes found on slide: ' + shapes.length);

 // Shuffle shapes to ensure random distribution
 shuffleArray(shapes);

 // Step 1: Determine group size (groups of 4)
 const groupSize = 4;
 const numGroups = Math.ceil(shapes.length / groupSize); // Total number of groups needed

 var groups = new Array(numGroups).fill(null).map(() => []);

 // Step 2: Place each selected student in their own group if possible
 selectedStudents.forEach(function (student, index) {
   var studentShape = shapes.find(function (shape) {
     return shape.getText().asString().trim() === student;
   });

   if (studentShape) {
     // Place the selected student into the group
     groups[index % numGroups].push(studentShape);

     // Debug: Alert the student being placed into a group
     SlidesApp.getUi().alert('Placed student: ' + student);

     // Remove the student from the shapes array to avoid redistributing them
     shapes = shapes.filter(function (shape) {
       return shape !== studentShape;
     });
   } else {
     // Debug: Alert if the student was not found on the slide
     SlidesApp.getUi().alert('Could not find shape for student: ' + student);
   }
 });

 // Step 3: Distribute remaining students randomly among the groups
 shapes.forEach(function (shape) {
   var groupWithSpace = groups.find(group => group.length < groupSize); // Find a group with space
   if (groupWithSpace) {
     groupWithSpace.push(shape);
   }
 });

 // Debug: Alert the total groups created and their sizes
 SlidesApp.getUi().alert('Total groups created: ' + groups.length + '. Group sizes: ' + groups.map(g => g.length).join(', '));

 // Step 4: Arrange the groups on the slide
 arrangeGroupsOnSlide(groups);

 // Debug: Alert when groups are successfully arranged
 SlidesApp.getUi().alert('Groups successfully arranged on the slide.');
}

function arrangeGroupsOnSlide(groups) {
 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 groups.forEach(function (group) {
   var color = getRandomPastelColor();

   // Shuffle the order within each group
   shuffleArray(group);

   if (group.length === 4) {
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition).getFill().setSolidFill(color);
     group[3].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
   } else if (group.length === 3) {
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH / 2).getFill().setSolidFill(color);
   } else if (group.length === 5) {
     // Arrange in 2x2 square with one student at the bottom middle
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition).getFill().setSolidFill(color);
     group[3].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[4].setTop(yPosition + 1.5 * RECT_HEIGHT).setLeft(xPosition + 0.5 * RECT_WIDTH).getFill().setSolidFill(color);
   }

   // Adjust positions for the next group
   groupCounter++;
   if (groupCounter % 3 === 0) {
     yPosition = 20;
     xPosition += 2 * RECT_WIDTH + GAP_BETWEEN_GROUPS + 20;
   } else {
     yPosition += 2 * RECT_HEIGHT + GAP_BETWEEN_GROUPS + 10;
   }
 });

 // Debug: Alert when groups are successfully arranged
 SlidesApp.getUi().alert('Groups successfully arranged.');
}

function groupAroundSelectedStudentsTest() {
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();

 // Hard-coded list of selected students (e.g., group leaders)
 var selectedStudents = ['Liz', 'Coco', 'Sophie', 'Ashley', 'Lesley'];

 // Log the selected students
 Logger.log('Selected Students: ' + selectedStudents.join(', '));

 // Get the class label shape ID from the speaker notes
 var notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 var match = notesText.match(/CLASS_LABEL_ID: (\w+)/);
 var classLabelId = match ? match[1] : null;

 var shapes = slide.getShapes();

 // Filter out the class label shape
 shapes = shapes.filter(function (shape) {
   return shape.getObjectId() !== classLabelId;
 });

 // Log the number of student shapes found on the slide
 Logger.log('Number of student shapes found on slide: ' + shapes.length);

 // Shuffle shapes to ensure random distribution
 shuffleArray(shapes);

 // Step 1: Determine group size (groups of 4)
 const groupSize = 4;
 const numGroups = Math.ceil(shapes.length / groupSize); // Total number of groups needed

 var groups = new Array(numGroups).fill(null).map(() => []);

 // Step 2: Place each selected student in their own group if possible
 selectedStudents.forEach(function (student, index) {
   var studentShape = shapes.find(function (shape) {
     return shape.getText().asString().trim() === student;
   });

   if (studentShape) {
     // Place the selected student into the group
     groups[index % numGroups].push(studentShape);

     // Log the student being placed into a group
     Logger.log('Placed student: ' + student);

     // Remove the student from the shapes array to avoid redistributing them
     shapes = shapes.filter(function (shape) {
       return shape !== studentShape;
     });
   } else {
     // Log if the student was not found on the slide
     Logger.log('Could not find shape for student: ' + student);
   }
 });

 // Step 3: Distribute remaining students randomly among the groups
 shapes.forEach(function (shape) {
   var groupWithSpace = groups.find(group => group.length < groupSize); // Find a group with space
   if (groupWithSpace) {
     groupWithSpace.push(shape);
   }
 });

 // Log the total groups created and their sizes
 Logger.log('Total groups created: ' + groups.length + '. Group sizes: ' + groups.map(g => g.length).join(', '));

 // Step 4: Arrange the groups on the slide
 arrangeGroupsOnSlide(groups);

 // Log when groups are successfully arranged
 Logger.log('Groups successfully arranged on the slide.');
}

function arrangeGroupsOnSlide(groups) {
 var xPosition = 20;
 var yPosition = 20;
 var groupCounter = 0;

 groups.forEach(function (group) {
   var color = getRandomPastelColor();

   // Shuffle the order within each group
   shuffleArray(group);

   if (group.length === 4) {
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition).getFill().setSolidFill(color);
     group[3].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
   } else if (group.length === 3) {
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH / 2).getFill().setSolidFill(color);
   } else if (group.length === 5) {
     // Arrange in 2x2 square with one student at the bottom middle
     group[0].setTop(yPosition).setLeft(xPosition).getFill().setSolidFill(color);
     group[1].setTop(yPosition).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[2].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition).getFill().setSolidFill(color);
     group[3].setTop(yPosition + RECT_HEIGHT).setLeft(xPosition + RECT_WIDTH).getFill().setSolidFill(color);
     group[4].setTop(yPosition + 1.5 * RECT_HEIGHT).setLeft(xPosition + 0.5 * RECT_WIDTH).getFill().setSolidFill(color);
   }

   // Adjust positions for the next group
   groupCounter++;
   if (groupCounter % 3 === 0) {
     yPosition = 20;
     xPosition += 2 * RECT_WIDTH + GAP_BETWEEN_GROUPS + 20;
   } else {
     yPosition += 2 * RECT_HEIGHT + GAP_BETWEEN_GROUPS + 10;
   }
 });

 // Log when groups are successfully arranged
 Logger.log('Groups arranged successfully.');
}


function arrangeInCircle() {
 const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 const scriptProperties = PropertiesService.getScriptProperties();
 const notesText = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
 const classLabelId = notesText.match(/CLASS_LABEL_ID: (\w+)/)?.[1];
 const resizedShapes = JSON.parse(scriptProperties.getProperty('resizedShapes') || '{}');

 let shapes = slide.getShapes().filter(shape => shape.getObjectId() !== classLabelId);

 shapes.forEach(shape => {
   if (!resizedShapes[shape.getObjectId()]) {
     shape.setWidth(shape.getWidth() * 0.85);
     shape.setHeight(shape.getHeight() * 0.85);
     resizedShapes[shape.getObjectId()] = true;
   }
   const textStyle = shape.getText().getTextStyle();
   textStyle.setBold(false);
   textStyle.setFontSize(11);  // Set the font size to 11
   shape.getFill().setSolidFill("#FFFFFF");
 });

 scriptProperties.setProperty('resizedShapes', JSON.stringify(resizedShapes));

 const numStudents = shapes.length;
 const [centerX, centerY] = [360, 240]; //centering horizontal and vertical on slide
 let [radiusX, radiusY] = [300, 205];  // Adjusted radius to reduce horizontal and vertical space between rectangles

 for (let attempts = 0; attempts < 100; attempts++) {
   shuffleArray(shapes);
   if (!shapes.some((shape, i) => isBadPair(shape, shapes[(i + 1) % shapes.length]))) break;
   if (attempts === 99) {
     SlidesApp.getUi().alert('Warning', 'Impossible to avoid bad pairings', Ui.ButtonSet.OK);
     return;
   }
 }

 shapes.forEach((shape, i) => {
   const angle = (2 * Math.PI / numStudents) * i;
   const x = centerX + radiusX * Math.cos(angle) - shape.getWidth() / 2;
   const y = centerY + radiusY * Math.sin(angle) - shape.getHeight() / 2;
   shape.setLeft(x).setTop(y);
 });

 initializeAvailableColors();
 shapes.forEach(shape => shape.getFill().setSolidFill(getRandomPastelColor()));

 const groupItems = shapes.map(shape => ({ objectId: shape.getObjectId() }));
 slide.createGroup(groupItems);
}

function selectRandomStudent() {
 deselectStudent(); // Deselect previously selected student if any

 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 if (shapes.length == 0) {
   return; // No shapes to select.
 }

 var randomIndex = Math.floor(Math.random() * shapes.length);
 var randomShape = shapes[randomIndex];
 // Bring the selected shape to the front
 randomShape.bringToFront();

 // Save the properties of the selected shape and slide

 scriptProperties.setProperty("lastSelectedShapeId", randomShape.getObjectId());
 scriptProperties.setProperty("lastSelectedFillColor", randomShape.getFill().getSolidFill().getColor().asRgbColor().asHexString());
 scriptProperties.setProperty("lastSelectedSlideId", slide.getObjectId());  // Save the slide ID

 randomShape.getFill().setSolidFill('#39FF14');  // Set fill color to indicate selection

 if (randomShape.getText()) {
   randomShape.getText().getTextStyle().setForegroundColor('#000000').setBold(true);
 }

 // Resize the shape
 var width = randomShape.getWidth();
 var height = randomShape.getHeight();
 randomShape.setWidth(width * 1.2);
 randomShape.setHeight(height * 1.2);
}

function deselectStudent() {

 var lastSelectedShapeId = scriptProperties.getProperty("lastSelectedShapeId");
 var lastSelectedFillColor = scriptProperties.getProperty("lastSelectedFillColor");
 var lastSelectedSlideId = scriptProperties.getProperty("lastSelectedSlideId");

 if (lastSelectedShapeId && lastSelectedFillColor && lastSelectedSlideId) {
   var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
   var currentSlideId = currentSlide.getObjectId();

   // If the slide has changed, load the last selected slide
   if (currentSlideId !== lastSelectedSlideId) {
     currentSlide = SlidesApp.getActivePresentation().getSlideById(lastSelectedSlideId);
   }

   var shapes = currentSlide.getShapes();
   var lastSelectedShape = null;

   // Search by shape ID first
   for (var i = 0; i < shapes.length; i++) {
     if (shapes[i].getObjectId() === lastSelectedShapeId) {
       lastSelectedShape = shapes[i];
       break;
     }
   }

   // If we can't find the shape by ID, look for it by its fill color
   if (!lastSelectedShape) {
     for (var i = 0; i < shapes.length; i++) {
       var shapeFillColor = shapes[i].getFill().getSolidFill().getColor().asRgbColor().asHexString();
       if (shapeFillColor === '#39FF14') {  // The highlight color from selectRandomStudent
         lastSelectedShape = shapes[i];
         break;
       }
     }
   }

   if (lastSelectedShape) {
     // Reset the fill color to its original
     lastSelectedShape.getFill().setSolidFill(lastSelectedFillColor);

     if (lastSelectedShape.getText()) {
       lastSelectedShape.getText().getTextStyle().setForegroundColor('#000000').setBold(false);  // Assuming default font color is black
     }

     // Reset the size
     var width = lastSelectedShape.getWidth();
     var height = lastSelectedShape.getHeight();
     lastSelectedShape.setWidth(width / 1.2);
     lastSelectedShape.setHeight(height / 1.2);

     // Clear the properties
     scriptProperties.deleteProperty("lastSelectedShapeId");
     scriptProperties.deleteProperty("lastSelectedFillColor");
     scriptProperties.deleteProperty("lastSelectedSlideId");
   }
 }
}



function shuffleArray(array) {
 for (let i = array.length - 1; i > 0; i--) {
   const j = Math.floor(Math.random() * (i + 1));
   [array[i], array[j]] = [array[j], array[i]]; // Swap elements
 }
}

function resizeAllRectangles() {
 var slides = SlidesApp.getActivePresentation().getSlides();

 for (var i = 0; i < slides.length; i++) {
   var shapes = slides[i].getShapes();

   for (var j = 0; j < shapes.length; j++) {
     var shape = shapes[j];

     // Check if the shape type is RECTANGLE
     if (shape.getShapeType() === SlidesApp.ShapeType.RECTANGLE) {
       shape.setWidth(.8 * 72);  // 1 inch is 72 points
       shape.setHeight(0.56 * 72);  //
     }
   }
 }
}


function resetTextPadding() {
 var slides = SlidesApp.getActivePresentation().getSlides();

 slides.forEach(function (slide) {
   var shapes = slide.getShapes();

   shapes.forEach(function (shape) {
     if (shape.getShapeType() === SlidesApp.ShapeType.RECTANGLE) {
       var textRange = shape.getText();
       var paragraphStyle = textRange.getParagraphStyle();

       paragraphStyle.setSpaceAbove(0);
       paragraphStyle.setSpaceBelow(0);
       paragraphStyle.setIndentStart(0);
       paragraphStyle.setIndentEnd(0);
     }
   });
 });
}

var availableColors = [];  // This will store the available colors during a function call

function initializeAvailableColors() {
 availableColors = [
   // Original Colors with English Names
   "#b7d8b7", // Soft Mint
   "#b7c9d8", // Light Sky Blue
   "#d8b7c9", // Pale Pink Purple
   "#d8d0b7", // Light Khaki
   "#c9b7d8", // Lavender
   "#d0d8b7", // Pale Lime Green
   "#b7d8c9", // Soft Aqua
   "#d8b7b7", // Soft Pink
   "#d8c9b7", // Pale Sand
   "#c9d8b7", // Light Green
   "#d5a6bd", // Muted Pink
   "#fff2cc", // Pale Yellow
   "#d9d2e9", // Pale Purple
   "#b6d7a8", // Light Green
   "#fce5cd", // Peach Cream
   "#e6b8af", // Soft Coral
   "#d0e0e3", // Pale Turquoise
   "#f4cccc", // Pastel Red
   "#ead1dc", // Pale Lavender Pink
   "#cfe2f3", // Soft Blue
   "#add8e6", // Light Blue
   "#f0e68c", // Khaki
   "#ffb6c1", // Light Pink
   "#d8bfd8", // Thistle
   "#dda0dd", // Plum
   "#ffe4e1", // Misty Rose
   "#ffebcd", // Blanched Almond
   "#fafad2", // Light Goldenrod Yellow
   "#ffe4b5", // Moccasin
   "#ffdead", // Navajo White
   "#f0e5c9", // Cream
   "#faf0e6", // Linen
   "#e6e6fa", // Lavender
   "#fff5ee", // Seashell
   "#f5f5dc", // Beige
   "#fdfd96", // Pastel Yellow
   "#a4c2f4", // Soft Blue
   "#9fc5e8", // Sky Blue
   "#6d9eeb", // Periwinkle Blue
   "#c9daf8", // Lavender Blue
   "#76a5af", // Soft Teal
   "#92cddc", // Light Blue-Green
   "#b4a7d6", // Light Purple
   "#8e7cc3", // Lavender Purple
   "#6fa8dc", // Cornflower Blue
   "#8faabd",  // Slate Gray Blue
   "#e3eaa7", // Pastel Lime
   "#d9e4fc", // Pale Sky Blue
   "#c5e3f6", // Powder Blue
   "#f7fcb9", // Soft Lemon
   "#d8e4bc", // Pale Olive Green
   "#cad2c5", // Sage Green
   "#e4e4c5", // Light Moss
   "#b8e0a2", // Pale Spring Green
   "#a2c4c9", // Soft Cyan
   "#ace1af",  // Celadon Green
   "#FFB3BA", // Light Salmon Pink
   "#FFDFBA", // Light Peach
   "#FFFFBA", // Light Butter Yellow
   "#BAFFC9", // Mint Green
   "#BAE1FF", // Baby Blue
 ];
}




function getRandomPastelColor() {
 if (availableColors.length === 0) {
   initializeAvailableColors();
 }
 var randomIndex = Math.floor(Math.random() * availableColors.length);
 var selectedColor = availableColors[randomIndex];
 availableColors.splice(randomIndex, 1);  // Remove the chosen color from the available colors
 return selectedColor;
}


//Still not being used
function saveLayout(layoutName) {
 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
 var shapes = slide.getShapes();

 var layoutData = shapes.map(function (shape) {

   var fillColor = "No Color"; // Default
   try {
     var fill = shape.getFill();
     if (fill) {
       var solidFill = fill.getSolidFill();
       if (solidFill) {
         var color = solidFill.getColor();
         if (color && typeof color.asRgbHex === 'function') {
           fillColor = color.asRgbHex();
         }
       }
     }
   } catch (e) {
     console.log("Error in extracting color: " + e.message);
   }



   return {
     id: shape.getObjectId(),
     top: shape.getTop(),
     left: shape.getLeft(),
     width: shape.getWidth(),
     height: shape.getHeight(),
     fillColor: fillColor,
     text: shape.getText().asString()
   };
 });

 var layoutJSON = JSON.stringify(layoutData);

 try {
   // TODO: Define this function to actually save the layoutJSON to Google Drive
   saveToDrive(layoutName, layoutJSON);
 } catch (e) {
   console.log("Error in saving layout to Drive: " + e.message);
   var ui = SlidesApp.getUi();
   ui.alert('Could not save the layout. Please try again.');
 }
}

// TODO: Define this function
function saveToDrive(layoutName, layoutJSON) {
 // Code to save layoutJSON to Google Drive
}



function loadLayout(layoutName) {
 // For this example, let's say we have a function loadFromDrive() that loads it from Google Drive
 var layoutJSON = loadFromDrive(layoutName);

 var layoutData = JSON.parse(layoutJSON);

 var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();

 layoutData.forEach(function (data) {
   var shape = slide.getShapes().find(function (shape) {
     return shape.getObjectId() === data.id;
   });

   if (shape) {
     shape.setTop(data.top).setLeft(data.left).setWidth(data.width).setHeight(data.height);
     shape.getFill().setSolidFill(data.fillColor);
     shape.getText().setText(data.text);  // Be cautious with this if text shouldn't change
   }
 });
}