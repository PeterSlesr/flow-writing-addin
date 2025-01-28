let isFlowModeActive = false;

Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
    }
});

// Main toggle for flow mode
async function toggleFlowMode() {
    console.log("toggleFlowMode called");
    try {
        isFlowModeActive = !isFlowModeActive;
        console.log("Flow mode is now:", isFlowModeActive);
        
        if (isFlowModeActive) {
            // Move cursor to end initially
            await Word.run(async (context) => {
                const body = context.document.body;
                const range = body.getRange('End');
                range.select();
                await context.sync();
            });
            
            // Add key event listener
            document.addEventListener('keydown', handleKeyPress, true);
            
            // Update status
            const statusDiv = document.getElementById('status');
            if (statusDiv) {
                statusDiv.textContent = 'Flow Mode: ON';
                statusDiv.className = 'active';
            }
        } else {
            // Remove event listener
            document.removeEventListener('keydown', handleKeyPress, true);
            
            // Update status
            const statusDiv = document.getElementById('status');
            if (statusDiv) {
                statusDiv.textContent = 'Flow Mode: OFF';
                statusDiv.className = 'inactive';
            }
        }
    } catch (error) {
        console.error("Error in toggleFlowMode:", error);
    }
}

// Handle key events
async function handleKeyPress(e) {
    if (!isFlowModeActive) return;
    
    console.log("Key pressed:", e.key);
    
    // Prevent backspace and delete
    if (e.key === 'Backspace' || e.key === 'Delete') {
        console.log("Blocking delete/backspace");
        e.preventDefault();
        e.stopPropagation();
        return false;
    }
    
    // Handle regular typing
    if (e.key.length === 1) { // Single character keys
        console.log("Processing character:", e.key);
        e.preventDefault();
        e.stopPropagation();
        
        try {
            await Word.run(async (context) => {
                const body = context.document.body;
                
                // Always insert at the end
                body.insertText(e.key, Word.InsertLocation.end);
                
                // Move selection to end
                const range = body.getRange('End');
                range.select();
                
                await context.sync();
            });
        } catch (error) {
            console.error("Error in handleKeyPress:", error);
        }
        return false;
    }
}