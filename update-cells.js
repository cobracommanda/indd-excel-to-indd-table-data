function updateMultipleParagraphs(document, updates) {
    var results = []; // To store the results of each update
  
    // Ensure 'updates' is an array
    if (!(updates instanceof Array)) {
      return ["Invalid updates data: not an array"];
    }
  
    // Iterate over the array with a traditional for loop
    for (var i = 0; i < updates.length; i++) {
      var update = updates[i];
      for (var j = 0; j < update.paragraphContents.length; j++) {
        var paragraph = update.paragraphContents[j];
        var path = paragraph.path;
        var newContent = paragraph.content;
  
        // Regex to extract indices from the path
        var regex = /Table(\d+)Row(\d+)Cell(\d+)Paragraph(\d+)/;
        var match = path.match(regex);
  
        if (match) {
          // Extract indices from the path
          var tableIndex = parseInt(match[1], 10);
          var rowIndex = parseInt(match[2], 10);
          var cellIndex = parseInt(match[3], 10);
          var paragraphIndex = parseInt(match[4], 10);
  
          // Access the target paragraph
          try {
            var targetTable = document.textFrames
              .everyItem()
              .tables.item(tableIndex);
            var targetCell = targetTable.rows
              .item(rowIndex)
              .cells.item(cellIndex);
            var targetParagraph = targetCell.paragraphs.item(paragraphIndex);
  
            // Update the paragraph content
            targetParagraph.contents = newContent;
            results.push("Updated: " + path);
          } catch (error) {
            results.push(
              "Failed to update: " + path + " with error: " + error.message
            );
          }
        } else {
          results.push("Invalid path provided for: " + path);
        }
      }
    }
  
    return results;
  }
  
  // Example usage
  if (app.documents.length > 0) {
    var myDocument = app.activeDocument;
    var updates = [
      {
        rowSpan: 1,
        colSpan: 1,
        paragraphStyles: ["•Table_Hd"],
        paragraphContents: [
          { content: "Grease 1", path: "Table0Row0Cell1Paragraph0" },
        ],
      },
      {
        rowSpan: 1,
        colSpan: 1,
        paragraphStyles: ["•Table_Hd"],
        paragraphContents: [
          { content: "Grease 2", path: "Table0Row0Cell2Paragraph0" },
        ],
      },
      {
        rowSpan: 1,
        colSpan: 1,
        paragraphStyles: ["•Table_Hd"],
        paragraphContents: [
          { content: "Grease 3", path: "Table0Row0Cell3Paragraph0" },
        ],
      },
      {
        rowSpan: 1,
        colSpan: 1,
        paragraphStyles: ["•Table_Hd"],
        paragraphContents: [
          { content: "Grease 4", path: "Table0Row0Cell4Paragraph0" },
        ],
      },
    ];
  
    var result = updateMultipleParagraphs(myDocument, updates);
    // for (var k = 0; k < result.length; k++) {
    //   alert(result[k]);
    // }
  } else {
    alert("No documents are open.");
  }
  