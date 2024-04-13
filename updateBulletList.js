function findAndUpdateOrDeleteText(document, searchText, replaceText) {
    try {
      // Clear any existing find/change settings at the start
      app.findTextPreferences = NothingEnum.nothing;
      app.changeTextPreferences = NothingEnum.nothing;
  
      // Validate inputs
      if (!document || !searchText) {
        throw new Error("Invalid input parameters.");
      }
  
      // Set the text to find
      app.findTextPreferences.findWhat = searchText;
  
      // Perform the find operation to check if there are occurrences before proceeding
      var initialFind = document.findText();
      if (initialFind.length === 0) {
        alert("No occurrences found initially.");
        return; // Exit if no occurrences to avoid unnecessary processing
      }
  
      // Check if replaceText is provided and it's not an empty string (assuming empty means no update)
      if (replaceText !== null && replaceText !== "") {
        app.changeTextPreferences.changeTo = replaceText;
      }
  
      // Perform the find and replace/delete operation
      var foundTexts = document.findText();
      var changesMade = 0;
  
      if (foundTexts.length > 0) {
        for (var i = 0; i < foundTexts.length; i++) {
          if (replaceText !== null && replaceText !== "") {
            foundTexts[i].changeText(); // Replace the text
          } else {
            foundTexts[i].remove(); // Delete the text
          }
          changesMade++;
        }
       
      } 
    } catch (error) {
  
    } finally {
      // Reset the find/change preferences
      app.findTextPreferences = NothingEnum.nothing;
      app.changeTextPreferences = NothingEnum.nothing;
    }
  }
  
  var templates = ["BULLET 1", "BULLET 2", "BULLET 3", "BULLET 4"];
  
  var data = {
    replaceText: [
      "Interactions with the natural world bring up strong feelings and emotions in people.",
      "Nature’s beauty and encounters with nature are recurring themes in literature. Characters reveal themselves through their nature.",
      "Interactions with the natural world bring up strong feelings and emotions in people.",
      "Nature’s beauty and encounters with nature are recurring themes in literature. Characters reveal themselves through their nature.",
    ],
  };
  
  for (var i = 0; i < data.replaceText.length; i++) {
    findAndUpdateOrDeleteText(document, templates[i], data.replaceText[i]);
  }
  