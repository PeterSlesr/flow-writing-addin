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
        console.log("Flow mode switched to:", isFlowModeActive);
        
        if (isFlowModeActive) {
            await Word.run(async (context) => {
                const doc = context.document;
                const body = doc.body;
                
                // Load current content
                body.load("text");
                await context.sync();
                const currentText = body.text;
                
                // Create a content control
                const contentControl = body.insertContentControl();
                contentControl.insertText(currentText, Word.InsertLocation.start);
                
                await context.sync();
                
                // Move to end
                const range = body.getRange('End');
                range.select();
                
                await context.sync();
            });
            updateStatus(true);
        } else {
            await Word.run(async (context) => {
                const body = context.document.body;
                
                // Just remove any content controls
                const contentControls = body.contentControls;
                contentControls.load("items");
                await context.sync();
                
                contentControls.items.forEach((control) => {
                    control.delete(false); // false = preserve content
                });
                
                await context.sync();
            });
            updateStatus(false);
        }
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