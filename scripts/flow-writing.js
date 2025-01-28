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
            await Word.run(async (context) => {
                // Protect the document
                const doc = context.document;
                doc.protection.set({
                    type: 'readOnly',
                    exceptions: ['everyone']
                });
                
                // Move to end and set up for typing
                const range = doc.body.getRange('End');
                range.select();
                
                await context.sync();
            });
            
            // Start cursor control
            startCursorControl();
            
            // Add key handlers
            document.addEventListener('keydown', handleKeyPress, true);
            window.addEventListener('keydown', handleKeyPress, true);
            
            updateStatus(true);
        } else {
            await Word.run(async (context) => {
                // Remove protection
                context.document.protection.unset();
                await context.sync();
            });
            
            // Remove handlers
            document.removeEventListener('keydown', handleKeyPress, true);
            window.removeEventListener('keydown', handleKeyPress, true);
            
            stopCursorControl();
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

// Handle key events
async function handleKeyPress(e) {
    if (!isFlowModeActive) return;

    // Block these keys completely
    const blockedKeys = [
        'Backspace', 'Delete', 
        'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown',
        'Home', 'End', 'PageUp', 'PageDown'
    ];
    
    if (blockedKeys.includes(e.key)) {
        e.preventDefault();
        e.stopPropagation();
        return false;
    }
    
    // Handle typing
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        try {
            await Word.run(async (context) => {
                const doc = context.document;
                
                // Temporarily unprotect
                doc.protection.unset();
                await context.sync();
                
                // Insert text at end
                const body = doc.body;
                body.insertText(e.key, Word.InsertLocation.end);
                
                // Move cursor to end
                const range = body.getRange('End');
                range.select();
                
                // Reprotect
                doc.protection.set({
                    type: 'readOnly',
                    exceptions: ['everyone']
                });
                
                await context.sync();
            });
        } catch (error) {
            console.error("Error typing:", error);
        }
        return false;
    }
}

// Aggressively control cursor position
function startCursorControl() {
    if (cursorInterval) {
        clearInterval(cursorInterval);
    }
    
    cursorInterval = setInterval(async () => {
        if (isFlowModeActive) {
            try {
                await Word.run(async (context) => {
                    const range = context.document.body.getRange('End');
                    range.select();
                    await context.sync();
                });
            } catch (error) {
                console.error("Error in cursor control:", error);
                isFlowModeActive = false;
                stopCursorControl();
                updateStatus(false);
            }
        }
    }, 50);
}

function stopCursorControl() {
    if (cursorInterval) {
        clearInterval(cursorInterval);
        cursorInterval = null;
    }
}