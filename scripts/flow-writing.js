let isFlowModeActive = false;
let cursorInterval = null;

Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
        
        // Add global key event listeners at the document level
        document.addEventListener('keydown', handleKeyPress, true);
        document.addEventListener('keyup', handleKeyPress, true);
        document.addEventListener('keypress', handleKeyPress, true);
        
        // Also add to window level to catch any that bubble up
        window.addEventListener('keydown', handleKeyPress, true);
        window.addEventListener('keyup', handleKeyPress, true);
        window.addEventListener('keypress', handleKeyPress, true);
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
            updateStatus(true);
        } else {
            // Stop cursor control
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

// Handle ALL key events
function handleKeyPress(e) {
    if (!isFlowModeActive) return;

    // List of keys to always block in flow mode
    const blockedKeys = [
        'Backspace',
        'Delete',
        'ArrowLeft',
        'ArrowRight',
        'ArrowUp',
        'ArrowDown',
        'Home',
        'End',
        'PageUp',
        'PageDown'
    ];
    
    // If it's a blocked key, prevent it completely
    if (blockedKeys.includes(e.key)) {
        console.log("Blocking navigation key:", e.key);
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        return false;
    }
    
    // For regular typing, only handle on keydown
    if (e.type === 'keydown' && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        console.log("Processing character:", e.key);
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        
        // Insert the character at the end
        Word.run(async (context) => {
            try {
                const body = context.document.body;
                body.insertText(e.key, Word.InsertLocation.end);
                const range = body.getRange('End');
                range.select();
                await context.sync();
            } catch (error) {
                console.error("Error in handleKeyPress:", error);
                isFlowModeActive = false;
                stopCursorControl();
                updateStatus(false);
            }
        });
        
        return false;
    }
}