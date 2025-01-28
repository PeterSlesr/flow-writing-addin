let isFlowModeActive = false;
let selectionInterval = null;

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
            
            // Start monitoring selection
            startSelectionMonitoring();
            
            // Add key event listener
            document.addEventListener('keydown', handleKeyPress, true);
            
            // Update status
            updateStatus(true);
        } else {
            // Stop monitoring selection
            stopSelectionMonitoring();
            
            // Remove key event listener
            document.removeEventListener('keydown', handleKeyPress, true);
            
            // Update status
            updateStatus(false);
        }
    } catch (error) {
        console.error("Error in toggleFlowMode:", error);
    }
}

function updateStatus(active) {
    const statusDiv = document.getElementById('status');
    if (statusDiv) {
        statusDiv.textContent = active ? 'Flow Mode: ON' : 'Flow Mode: OFF';
        statusDiv.className = active ? 'active' : 'inactive';
    }
}

// Start monitoring selection changes
function startSelectionMonitoring() {
    if (selectionInterval) {
        clearInterval(selectionInterval);
    }
    
    selectionInterval = setInterval(async () => {
        if (isFlowModeActive) {
            await Word.run(async (context) => {
                const doc = context.document;
                const selection = doc.getSelection();
                const body = doc.body;
                
                selection.load('start');
                body.load('text');
                await context.sync();
                
                // If selection is not at the end, move it there
                if (selection.start < body.text.length) {
                    const range = body.getRange('End');
                    range.select();
                    await context.sync();
                }
            });
        }
    }, 100); // Check every 100ms
}

// Stop monitoring selection changes
function stopSelectionMonitoring() {
    if (selectionInterval) {
        clearInterval(selectionInterval);
        selectionInterval = null;
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