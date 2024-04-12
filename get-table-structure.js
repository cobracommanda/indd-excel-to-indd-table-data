var targetScriptLabel = "planner";

// Function to create a unique key for each cell
function createCellKey(tableIndex, rowIndex, cellIndex) {
    return 'Table' + tableIndex + 'Row' + rowIndex + 'Cell' + cellIndex;
}

// Enhanced JSON serialization function to ensure proper escaping
function toJSONString(obj) {
    var parts = [];
    var isList = (Object.prototype.toString.call(obj) === '[object Array]');

    for (var key in obj) {
        var value = obj[key];
        if (typeof value == "object" && value !== null) {
            if (isList) parts.push(toJSONString(value));
            else parts.push('"' + key + '": ' + toJSONString(value));
        } else {
            var strValue = value.toString();
            if (typeof value === "string") {
                // Escaping backslashes, quotes, and control characters
                strValue = strValue.replace(/\\/g, '\\\\')
                                   .replace(/"/g, '\\"')
                                   .replace(/\n/g, '\\n')
                                   .replace(/\r/g, '\\r')
                                   .replace(/\f/g, '\\f')
                                   .replace(/\t/g, '\\t');
                strValue = '"' + strValue + '"';
            }
            if (isList) parts.push(strValue);
            else parts.push('"' + key + '": ' + strValue);
        }
    }
    var json = parts.join(",");
    if (isList) return '[' + json + ']';  
    return '{' + json + '}'; 
}

if (app.documents.length > 0) {
    var myDocument = app.activeDocument;
    var cellsArray = []; // Array to hold detailed objects for each cell
    var cellCounter = 1; // Counter to help generate standardized content keys

    // Find text frames with the specified script label
    var frames = myDocument.textFrames.everyItem().getElements();
    var targetFrameFound = false;

    for (var i = 0; i < frames.length; i++) {
        if (frames[i].label == targetScriptLabel) {
            targetFrameFound = true;
            var myFrame = frames[i];

            if (myFrame.tables.length > 0) {
                for (var tableIndex = 0; tableIndex < myFrame.tables.length; tableIndex++) {
                    var myTable = myFrame.tables.item(tableIndex);

                    for (var rowIndex = 0; rowIndex < myTable.rows.length; rowIndex++) {
                        var myRow = myTable.rows.item(rowIndex);

                        for (var cellIndex = 0; cellIndex < myRow.cells.length; cellIndex++) {
                            var myCell = myRow.cells.item(cellIndex);
                            var paragraphStyles = [];
                            var paragraphContents = [];

                            // Collect all paragraph styles and contents from each paragraph in the cell
                            for (var p = 0; p < myCell.paragraphs.length; p++) {
                                var paragraph = myCell.paragraphs[p];
                                paragraphStyles.push(paragraph.appliedParagraphStyle.name);
                                paragraphContents.push(paragraph.contents);
                            }

                            // Create a detailed object for each cell
                            var cellObject = {
                                rowSpan: myCell.rowSpan,
                                colSpan: myCell.columnSpan,
                                paragraphStyles: paragraphStyles,
                                paragraphContents: paragraphContents // Storing paragraph contents
                            };

                            // Add the cell object to the array
                            cellsArray.push(cellObject);
                            cellCounter++; // Increment the counter for the next cell
                        }
                    }
                }
            } else {
                alert("The text frame with the script label '" + targetScriptLabel + "' does not contain any tables.");
            }
            break; // Stop once the target frame is found and processed
        }
    }

    if (!targetFrameFound) {
        alert("No text frame with the script label '" + targetScriptLabel + "' was found.");
    } else {
        // Serialize the array and log it
        var serializedData = toJSONString(cellsArray);
        
        alert(serializedData);
    }
} else {
    alert("No documents are open.");
}
