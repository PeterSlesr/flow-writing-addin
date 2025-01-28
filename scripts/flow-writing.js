let isFlowModeActive = false;
let cursorInterval = null;

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
            // Start aggressive cursor control
            startCursorControl();
            
            // Add key event listener
            document.addEventListener('keydown', handleKeyPress, true);
            
            // Update status
            updateStatus(true);
        } else {
            // Stop cursor control
            stopCursorControl();
            
            // Remove key event listener
            document.removeEventListener('keydown', handleKeyPress, true);
            
            // Update status
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

// Aggressively control cursor position
function startCursorControl() {
    if (cursorInterval) {
        clearInterval(cursorInterval);
    }
    
    // Check and correct cursor position very frequently
    cursorInterval = setInterval(async () => {
        if (isFlowModeActive) {
            try {
                await Word.run(async (context) => {
                    const document = context.document;
                    // Force selection to end
                    const range = document.body.getRange('End');
                    range.select();
                    await context.sync();
                });
            } catch (error) {
                console.error("Error in cursor control:", error);
                // If we get an error, assume we lost control and disable flow mode
                isFlowModeActive = false;
                stopCursorControl();
                updateStatus(false);
            }
        }
    }, 50); // Check every 50ms - very aggressive
}

function stopCursorControl() {
    if (cursorInterval) {
        clearInterval(cursorInterval);
        cursorInterval = null;
    }
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
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        console.log("Processing character:", e.key);
        e.preventDefault();
        e.stopPropagation();
        
        try {
            await Word.run(async (context) => {
                const body = context.document.body;
                
                // Always insert at the end
                body.insertText(e.key, Word.InsertLocation.end);
                
                // Force selection to end
                const range = body.getRange('End');
                range.select();
                
                await context.sync();
            });
        } catch (error) {
            console.error("Error in handleKeyPress:", error);
            // If we get an error, disable flow mode
            isFlowModeActive = false;
            stopCursorControl();
            updateStatus(false);
        }
        return false;
    }
}