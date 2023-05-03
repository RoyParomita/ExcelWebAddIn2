'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#search').click(searchArtwork);
            //setColor();
        });
    });

    async function searchArtwork() {
        await Excel.run(async (context) => {
            //const range = context.workbook.getSelectedRange();
            //range.format.fill.color = 'green';

            let searchString = document.getElementById("search-string").value;
            let clearValues = [
                ['', ''],
                ['', '']
            ];
            let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:B2");
            range.values = clearValues;
            
            axios.get("https://collectionapi.metmuseum.org/public/collection/v1/objects/" + searchString)
                .then(function (response) {
                    // handle success
                    console.log(response);
                    let title = response.data.title;
                    let artist = response.data.artistDisplayName;

                    let columnHeader = [
                        ['Title', 'Artist'],
                    ];
                    range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:B1");
                    range.values = columnHeader
                    range.format.fill.color = 'gray';

                    console.log(title);
                    range = context.workbook.worksheets.getActiveWorksheet().getRange("A2");
                    range.values = [[title]];
                    range.format.autofitColumns();

                    console.log(artist);
                    range = context.workbook.worksheets.getActiveWorksheet().getRange("B2");
                    range.values = [[artist]];
                    range.format.autofitColumns();

                    context.sync();
                })
                .catch(function (error) {
                    // handle error
                    console.log(error);
                    
                    range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
                    range.values = [[error.message]];

                    context.sync();
                })
                .finally(function () {
                    // always executed
                });

            await context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();
