let isFlowModeActive = false;
let cursorInterval = null;

Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
        
        // Set up Word document event handlers
        setupWordHandlers();
    }
});

async function setupWordHandlers() {
    try {
        await Word.run(async (context) => {
            // Handle document events
            context.document.onContentControlEntered.add(handleWordKeyEvent);
            context.document.onSelectionChanged.add(handleSelectionChange);
            
            // Bind to the document body
            const body = context.document.body;
            body.onKeyDown.add(handleWordKeyEvent);
            
            await context.sync();
        });
    } catch (error) {
        console.error("Error setting up Word handlers:", error);
    }
}

async function handleWordKeyEvent(event) {
    if (!isFlowModeActive) return;
    
    console.log("Word key event:", event);
    
    // Block navigation keys
    try {
        await Word.run(async (context) => {
            // Move cursor to end
            const range = context.document.body.getRange('End');
            range.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error in handleWordKeyEvent:", error);
    }
}

async function handleSelectionChange(event) {
    if (!isFlowModeActive) return;
    
    try {
        await Word.run(async (context) => {
            // Force cursor to end
            const range = context.document.body.getRange('End');
            range.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error in handleSelectionChange:", error);
    }
}

// Main toggle for flow mode
async function toggleFlowMode() {
    console.log("toggleFlowMode called");
    try {
        isFlowModeActive = !isFlowModeActive;
        console.log("Flow mode is now:", isFlowModeActive);
        
        if (isFlowModeActive) {
            await Word.run(async (context) => {
                // Set document to read-only when flow mode is on
                context.document.body.style.readOnly = true;
                
                // But allow insertions at the end
                const range = context.document.body.getRange('End');
                range.select();
                range.style.readOnly = false;
                
                await context.sync();
            });
            
            // Start cursor control
            startCursorControl();
            updateStatus(true);
        } else {
            await Word.run(async (context) => {
                // Remove read-only when flow mode is off
                context.document.body.style.readOnly = false;
                await context.sync();
            });
            
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
    
    cursorInterval = setInterval(async () => {
        if (isFlowModeActive) {
            try {
                await Word.run(async (context) => {
                    // Keep document in read-only mode except for end
                    context.document.body.style.readOnly = true;
                    const range = context.document.body.getRange('End');
                    range.select();
                    range.style.readOnly = false;
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