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
            await moveToEnd();
            
            // Add event listeners
            document.addEventListener('keydown', handleKeyPress, true);
            document.addEventListener('selectionchange', handleSelectionChange, true);
            document.addEventListener('click', handleClick, true);
            
            // Update status
            const statusDiv = document.getElementById('status');
            if (statusDiv) {
                statusDiv.textContent = 'Flow Mode: ON';
                statusDiv.className = 'active';
            }
        } else {
            // Remove event listeners
            document.removeEventListener('keydown', handleKeyPress, true);
            document.removeEventListener('selectionchange', handleSelectionChange, true);
            document.removeEventListener('click', handleClick, true);
            
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

// Move cursor to end of document
async function moveToEnd() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const range = body.getRange('End');
            range.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error moving to end:", error);
    }
}

// Handle clicks
function handleClick(e) {
    if (!isFlowModeActive) return;
    
    // Prevent click from repositioning cursor
    e.preventDefault();
    e.stopPropagation();
    moveToEnd();
    return false;
}

// Handle selection changes
function handleSelectionChange(e) {
    if (!isFlowModeActive) return;
    
    // Move back to end when selection changes
    moveToEnd();
}

// Handle key events
async function handleKeyPress(e) {
    if (!isFlowModeActive) return;
    
    console.log("Key pressed:", e.key);
    
    // Prevent navigation keys
    if (e.key === 'Backspace' || 
        e.key === 'Delete' || 
        e.key === 'ArrowLeft' || 
        e.key === 'ArrowRight' ||
        e.key === 'Home' ||
        e.key === 'End') {
        console.log("Blocking navigation key");
        e.preventDefault();
        e.stopPropagation();
        return false;
    }
    
    // Handle regular typing
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { // Single character keys without modifiers
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