let isFlowModeActive = false;
let flowControl = null;

Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
    }
});

async function toggleFlowMode() {
    console.log("toggleFlowMode called");
    try {
        isFlowModeActive = !isFlowModeActive;
        console.log("Flow mode is now:", isFlowModeActive);
        
        await Word.run(async (context) => {
            const doc = context.document;
            
            if (isFlowModeActive) {
                // Store current content
                const body = doc.body;
                body.load("text");
                await context.sync();
                const currentText = body.text;
                
                // Clear the document
                body.clear();
                await context.sync();
                
                // Create two content controls
                // One for existing text (locked)
                const existingTextControl = body.insertContentControl();
                existingTextControl.insertText(currentText, Word.InsertLocation.start);
                existingTextControl.cannotDelete = true;
                existingTextControl.cannotEdit = true;
                existingTextControl.appearance = "Hidden";  // Hide the boundaries
                
                // One for new text (editable only at end)
                flowControl = body.insertContentControl();
                flowControl.cannotDelete = true;
                flowControl.appearance = "Hidden";
                
                // Move cursor to end
                const range = body.getRange('End');
                range.select();
                
                await context.sync();
                
                updateStatus(true);
            } else {
                // Combine content and remove controls
                const body = doc.body;
                body.load("text");
                await context.sync();
                
                const text = body.text;
                body.clear();
                body.insertText(text, Word.InsertLocation.start);
                
                flowControl = null;
                
                await context.sync();
                
                updateStatus(false);
            }
        });
    } catch (error) {
        console.error("Error in toggleFlowMode:", error);
        isFlowModeActive = false;
        updateStatus(false);
    }
}

function updateStatus(active) {
    const statusDiv = document.getElementById('status');
    if (statusDiv) {
        statusDiv.textContent = active ? 'Flow Mode: ON' : 'Flow Mode: OFF';
        statusDiv.className = active ? 'active' : 'inactive';
    }
}