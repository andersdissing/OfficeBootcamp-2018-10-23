(function() {
    "use strict";

    Office.initialize = function() {
        $(document).ready(function() {

            // TODO1: Assign handler to the OK button.
            $('#ok-button').click(sendStringToParentPage);
        });
    }

    // TODO2: Create the OK button handler
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
}());