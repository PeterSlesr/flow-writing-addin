// Word Flow Writing Plugin
Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
        document.getElementById('toggleVisibility').onclick = toggleTextVisibility;
    }
});

let isFlowModeActive = false;
let isTextHidden = false;
let originalContent = '';

// Main toggle for flow mode
async function toggleFlowMode() {
    console.log("toggleFlowMode called");
    isFlowModeActive = !isFlowModeActive;
    console.log("Flow mode is now:", isFlowModeActive);
    
    try {
        if (isFlowModeActive) {
            console.log("Activating flow mode");
            // Store original content
            await Word.run(async (context) => {
                console.log("Starting Word.run");
                const body = context.document.body;
                body.load("text");
                await context.sync();
                originalContent = body.text;
                console.log("Original content stored");
            });
            
            // Add key event listeners
            console.log("Adding key event listener");
            document.addEventListener('keydown', handleKeyPress);
            
            // Update status
            const statusDiv = document.getElementById('status');
            if (statusDiv) {
                statusDiv.textContent = 'Flow Mode: ON';
                statusDiv.className = 'active';
            }
        } else {
            console.log("Deactivating flow mode");
            // Remove event listeners
            document.removeEventListener('keydown', handleKeyPress);
            
            // Reset visibility if needed
            if (isTextHidden) {
                await toggleTextVisibility();
            }
            
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
    console.log("Key pressed:", e.key);
    if (isFlowModeActive) {
        // Prevent backspace and delete
        if (e.key === 'Backspace' || e.key === 'Delete') {
            console.log("Blocking delete/backspace");
            e.preventDefault();
            return;
        }
        
        // Handle regular typing
        if (e.key.length === 1) { // Single character keys
            console.log("Processing character:", e.key);
            e.preventDefault();
            
            try {
                await Word.run(async (context) => {
                    const body = context.document.body;
                    
                    // Always insert at the end
                    body.insertText(e.key, Word.InsertLocation.end);
                    
                    // If hiding previous text is active
                    if (isTextHidden) {
                        // Make all text white except last character
                        const range = body.getRange();
                        range.font.color = 'white';
                        
                        const lastChar = body.getRange(body.length - 1, body.length);
                        lastChar.font.color = 'black';
                    }
                    
                    await context.sync();
                });
            } catch (error) {
                console.error("Error in handleKeyPress:", error);
            }
        }
    }
}

// Toggle text visibility
async function toggleTextVisibility() {
    console.log("toggleTextVisibility called");
    if (!isFlowModeActive) {
        console.log("Flow mode not active, ignoring visibility toggle");
        return;
    }
    
    try {
        isTextHidden = !isTextHidden;
        console.log("Text hidden:", isTextHidden);
        
        await Word.run(async (context) => {
            const body = context.document.body;
            
            if (isTextHidden) {
                // Make all text white except last character
                body.font.color = 'white';
                if (body.length > 0) {
                    const lastChar = body.getRange(body.length - 1, body.length);
                    lastChar.font.color = 'black';
                }
            } else {
                // Make all text visible
                body.font.color = 'black';
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error("Error in toggleTextVisibility:", error);
    }
}