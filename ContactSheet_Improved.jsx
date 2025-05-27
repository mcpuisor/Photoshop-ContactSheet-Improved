app.bringToFront();

// Save original ruler units so we can restore later
var originalRulerUnits = app.preferences.rulerUnits;
app.preferences.rulerUnits = Units.PIXELS;

function main() {
    // Prompt the user for column count (allowed: 2,3,4,5,6)
    var columnInput = prompt("Enter number of columns (2, 3, 4, 5 or 6):", "3");
    var columnCount = parseInt(columnInput, 10);
    if (isNaN(columnCount) || columnCount < 2 || columnCount > 6) {
        alert("Invalid column count. Please run the script again and choose one of the following values: 2, 3, 4, 5 or 6.");
        return;
    }
    
    // Ask user to select a folder containing image files
    var inputFolder = Folder.selectDialog("Select the folder with images to use");
    if (inputFolder == null) {
        return; // cancelled
    }
    // Filter for common image file extensions (adjust as needed)
    var fileList = inputFolder.getFiles(/\.(jpg|jpeg|png|tif|tiff|bmp)$/i);
    if (fileList.length == 0) {
        alert("No image files found in the selected folder.");
        return;
    }
    
    // Determine if images are primarily horizontal or vertical
    var isHorizontal = true;
    if (fileList.length > 0) {
        var firstImage = fileList[0];
        var img = new File(firstImage);
        
        // Get image dimensions
        var imgInfo = app.open(img);
        isHorizontal = imgInfo.width > imgInfo.height;
        imgInfo.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Calculate row count based on document dimensions and image orientation
    var rowCount;
    if (isHorizontal) {
        // For landscape documents, use more rows for horizontal images
        rowCount = Math.ceil(6 * (isHorizontal ? 0.8 : 1.2));
    } else {
        // For portrait documents, use more rows for vertical images
        rowCount = Math.ceil(6 * (isHorizontal ? 1.2 : 0.8));
    }
    
    // Ensure we have at least 2 rows
    rowCount = Math.max(2, rowCount);
    
    // Define page size based on orientation
    var resolution = 300;
    if (isHorizontal) {
        // Landscape orientation (wide)
        var docWidth = 3508;
        var docHeight = 2480;
    } else {
        // Portrait orientation (tall)
        var docWidth = 2480;
        var docHeight = 3508;
    }
    
    // Define margin and gap between images (both 25 px)
    var margin = 25;
    var gap = 25;
    
    // Calculate effective cell dimensions considering margins and gaps
    var cellWidth = (docWidth - 2 * margin - (columnCount - 1) * gap) / columnCount;
    var cellHeight = (docHeight - 2 * margin - (rowCount - 1) * gap) / rowCount;
    
    var imageIndex = 0;
    var pageCount = 0;
    var imagesPerPage = columnCount * rowCount;
    
    // Create as many pages as needed
    while (imageIndex < fileList.length) {
        pageCount++;
        var contactDoc = app.documents.add(docWidth, docHeight, resolution, "Custom Contact Sheet - Page " + pageCount, NewDocumentMode.RGB, DocumentFill.WHITE);
        
        // Loop through grid cells in this page
        for (var row = 0; row < rowCount; row++) {
            for (var col = 0; col < columnCount; col++) {
                if (imageIndex >= fileList.length) break; // no more images
                
                var imageFile = fileList[imageIndex];
                imageIndex++;
                
                // --- Place the image as a smart object using the "Place" command ---
                var idPlc = charIDToTypeID("Plc ");
                var descPlc = new ActionDescriptor();
                descPlc.putPath(charIDToTypeID("null"), new File(imageFile));
                descPlc.putEnumerated(charIDToTypeID("FTcs"), charIDToTypeID("QCSt"), charIDToTypeID("Qcsa"));
                
                // Calculate the desired center position for the cell
                var centerX = margin + col * (cellWidth + gap) + cellWidth / 2;
                var centerY = margin + row * (cellHeight + gap) + cellHeight / 2;
                var descOffset = new ActionDescriptor();
                descOffset.putUnitDouble(charIDToTypeID('Hrzn'), charIDToTypeID('#Pxl'), centerX);
                descOffset.putUnitDouble(charIDToTypeID('Vrtc'), charIDToTypeID('#Pxl'), centerY);
                descPlc.putObject(charIDToTypeID("Ofst"), charIDToTypeID("Ofst"), descOffset);
                executeAction(idPlc, descPlc, DialogModes.NO);
                
                // After placement, the placed layer is active.
                var placedLayer = contactDoc.activeLayer;
                
                // --- Resize the image to fit within the cell ---
                var bounds = placedLayer.bounds;
                var layerWidth = bounds[2].value - bounds[0].value;
                var layerHeight = bounds[3].value - bounds[1].value;
                var scaleFactor = Math.min(cellWidth / layerWidth, cellHeight / layerHeight) * 100;
                placedLayer.resize(scaleFactor, scaleFactor, AnchorPosition.MIDDLECENTER);
                
                // --- Center the layer within its cell ---
                bounds = placedLayer.bounds;
                var currentCenterX = (bounds[0].value + bounds[2].value) / 2;
                var currentCenterY = (bounds[1].value + bounds[3].value) / 2;
                var deltaX = centerX - currentCenterX;
                var deltaY = centerY - currentCenterY;
                placedLayer.translate(deltaX, deltaY);
            }
            if (imageIndex >= fileList.length) break;
        }
    }
    
    alert("Custom contact sheet created with " + pageCount + " page(s) using a " + columnCount + "x" + rowCount + " grid layout.");
}

// Run the main function
main();

// Restore original ruler units
app.preferences.rulerUnits = originalRulerUnits;
