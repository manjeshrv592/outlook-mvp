/* global Office */

Office.onReady(() => {
  console.log("Commands loaded successfully!");
});

// Dummy button click handler
function dummyButtonClicked(event) {
  console.log("Dummy button was clicked!");
  
  // Show a notification (optional)
  Office.context.mailbox.item.notificationMessages.addAsync("dummyNotification", {
    type: "informationalMessage",
    message: "MVP Test: Button clicked successfully! 🎉",
    icon: "Icon.80x80",
    persistent: false
  });
  
  // Signal that the command is complete
  event.completed();
}

// Register the function
Office.actions.associate("dummyButtonClicked", dummyButtonClicked);
