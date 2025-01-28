// Word Flow Writing Plugin
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('toggleFlow').onclick = toggleFlowMode;
        document.getElementById('toggleVisibility').onclick = toggleTextVisibility;
    }
});

let isFlowModeActive = false;
let isTextHidden = false;
let originalContent = '';

// Main toggle for flow mode
async function toggleFlowMode() {
    isFlowModeActive = !isFlowModeActive;
    
    if (isFlowModeActive) {
        // Store original content
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            originalContent = body.text;
        });
        
        // Disable ribbon commands
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "TabHome",
                    enabled: false
                },
                {
                    id: "TabInsert",
                    enabled: false
                },
                // Add other tabs as needed
            ]
        });
        
        // Add key event listeners
        document.addEventListener('keydown', handleKeyPress);
    } else {
        // Re-enable ribbon
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "TabHome",
                    enabled: true
                },
                {
                    id: "TabInsert",
                    enabled: true
                }
            ]
        });
        
        // Remove event listeners
        document.removeEventListener('keydown', handleKeyPress);
        
        // Reset visibility if needed
        if (isTextHidden) {
            toggleTextVisibility();
        }
    }
}

// Handle key events
async function handleKeyPress(e) {
    if (isFlowModeActive) {
        // Prevent backspace and delete
        if (e.key === 'Backspace' || e.key === 'Delete') {
            e.preventDefault();
            return;
        }
        
        // Handle regular typing
        if (e.key.length === 1) { // Single character keys
            e.preventDefault();
            
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
        }
    }
}

// Toggle text visibility
async function toggleTextVisibility() {
    if (!isFlowModeActive) return;
    
    isTextHidden = !isTextHidden;
    
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
}

// Add to Word ribbon
Office.addin.onRibbonLoad = function(event) {
    // Create custom ribbon group
    const customGroup = {
        id: "flowWritingGroup",
        label: "Flow Writing",
        controls: [
            {
                type: "button",
                id: "toggleFlow",
                label: "Toggle Flow Mode",
                image: "flow-icon", // You'll need to provide this
                onClick: toggleFlowMode
            },
            {
                type: "button",
                id: "toggleVisibility",
                label: "Toggle Text Visibility",
                image: "visibility-icon", // You'll need to provide this
                onClick: toggleTextVisibility
            }
        ]
    };
    
    event.ribbon.tabs.push(customGroup);
};