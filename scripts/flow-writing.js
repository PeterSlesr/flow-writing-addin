let isFlowModeActive = false;
let isTextHidden = false;
let originalContent = '';

Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Word) {
        console.log("Word detected, setting up button handlers");
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
        document.getElementById('toggleVisibility').onclick = toggleTextVisibility;
        
        // Add event listener to Word document
        setupWordEventHandlers();
    }
});

async function setupWordEventHandlers() {
    try {
        await Word.run(async (context) => {
            // Handle document changes
            context.document.onContentControlAdded.add(handleDocumentChange);
            context.document.onSelectionChanged.add(handleSelectionChange);
            await context.sync();
        });
    } catch (error) {
        console.error("Error setting up Word handlers:", error);
    }
}

// Main toggle for flow mode
async function toggleFlowMode() {
    console.log("toggleFlowMode called");
    try {
        isFlowModeActive = !isFlowModeActive;
        console.log("Flow mode is now:", isFlowModeActive);
        
        if (isFlowModeActive) {
            // Store original content
            await Word.run(async (context) => {
                console.log("Starting Word.run");
                const body = context.document.body;
                body.load("text");
                await context.sync();
                originalContent = body.text;
                console.log("Original content stored");
                
                // Move cursor to end
                const range = body.getRange('End');
                range.select();
                await context.sync();
            });
            
            // Add key event listeners at both document and window level
            document.addEventListener('keydown', handleKeyPress, true);
            window.addEventListener('keydown', handleKeyPress, true);
            
            // Update status
            const statusDiv = document.getElementById('status');
            if (statusDiv) {
                statusDiv.textContent = 'Flow Mode: ON';
                statusDiv.className = 'active';
            }
        } else {
            // Remove event listeners
            document.removeEventListener('keydown', handleKeyPress, true);
            window.removeEventListener('keydown', handleKeyPress, true);
            
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
                
                // If hiding previous text is active
                if (isTextHidden) {
                    // Make all text white except last character
                    const fullRange = body.getRange();
                    fullRange.font.color = 'white';
                    
                    const lastChar = body.getRange(body.length - 1, body.length);
                    lastChar.font.color = 'black';
                }
                
                await context.sync();
            });
        } catch (error) {
            console.error("Error in handleKeyPress:", error);
        }
        return false;
    }
}

// Handle any document changes
async function handleDocumentChange(eventArgs) {
    if (!isFlowModeActive) return;
    
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const range = body.getRange('End');
            range.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error handling document change:", error);
    }
}

// Handle selection changes
async function handleSelectionChange(eventArgs) {
    if (!isFlowModeActive) return;
    
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const range = body.getRange('End');
            range.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error handling selection change:", error);
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