// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
    function getAllTable() {
        Excel.run(function (ctx) {
            var tables = ctx.workbook.tables;
            tables.load();
            return ctx.sync().then(function () {
                console.log("tables Count: " + tables.count);
                for (var i = 0; i < tables.items.length; i++) {
                    console.log(tables.items[i].name);
                }
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();