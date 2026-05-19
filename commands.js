/* DataPro Add-in — commands.js
   Ribbon button handler functions registered with Office.js
   MIT License */

Office.onReady(() => {});

// Called when user clicks "Smart Dashboard" ribbon button
function openDashboard(event) {
  event.completed();
}

// Called when user clicks "Heatmap" ribbon button
function openHeatmap(event) {
  event.completed();
}

// Called when user clicks "AI Assistant" ribbon button
function openAI(event) {
  event.completed();
}

// Called when user clicks "Data Cleaner" ribbon button
function openCleaner(event) {
  event.completed();
}

// Register all functions with Office
Office.actions.associate("openDashboard", openDashboard);
Office.actions.associate("openHeatmap", openHeatmap);
Office.actions.associate("openAI", openAI);
Office.actions.associate("openCleaner", openCleaner);
