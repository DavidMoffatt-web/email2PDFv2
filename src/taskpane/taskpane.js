document.addEventListener("DOMContentLoaded", function () {
    // Initialize the task pane
    document.getElementById("myButton").onclick = function () {
        // Handle button click
        Office.context.mailbox.item.body.getAsync("text", function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("output").innerText = result.value;
            } else {
                document.getElementById("output").innerText = "Error: " + result.error.message;
            }
        });
    };
});