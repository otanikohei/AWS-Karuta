#target illustrator

(function () {
    var ROW_TOLERANCE = 8; // 同じ行とみなす上下位置の差（pt）

    if (app.documents.length === 0) {
        alert("Illustratorドキュメントを開いてください。");
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    var textFrames = [];

    if (!selectedItems || selectedItems.length === 0) {
        alert("商品名を流し込むテキストオブジェクトを選択してください。");
        return;
    }

    for (var i = 0; i < selectedItems.length; i++) {
        collectTextFrames(selectedItems[i], textFrames);
    }

    if (textFrames.length === 0) {
        alert("選択範囲にテキストオブジェクトがありません。");
        return;
    }

    // 上から下、同じ行では左から右の順に並べる
    textFrames.sort(function (a, b) {
        var aBounds = a.geometricBounds; // [left, top, right, bottom]
        var bBounds = b.geometricBounds;
        var topDiff = bBounds[1] - aBounds[1];

        if (Math.abs(topDiff) > ROW_TOLERANCE) {
            return topDiff;
        }
        return aBounds[0] - bBounds[0];
    });

    var csvFile = File.openDialog("商品名が入ったCSVファイルを選択してください");
    if (!csvFile) {
        return;
    }

    csvFile.encoding = "UTF-8";
    if (!csvFile.open("r")) {
        alert("CSVファイルを開けませんでした。");
        return;
    }

    var csvText = csvFile.read();
    csvFile.close();

    // UTF-8 BOMを除去
    csvText = csvText.replace(/^\uFEFF/, "");

    var rows = parseCSV(csvText);
    if (rows.length === 0) {
        alert("CSVにデータがありません。");
        return;
    }

    // 先頭セルが一般的な見出しなら自動的に読み飛ばす
    var startRow = 0;
    var firstCell = trimText(rows[0][0]).toLowerCase();
    if (
        firstCell === "商品名" ||
        firstCell === "product_name" ||
        firstCell === "product name" ||
        firstCell === "name"
    ) {
        startRow = 1;
    }

    var productNames = [];
    for (var r = startRow; r < rows.length; r++) {
        // CSVの1列目を商品名として使用
        productNames.push(rows[r].length > 0 ? rows[r][0] : "");
    }

    if (productNames.length === 0) {
        alert("CSVの1列目に商品名がありません。");
        return;
    }

    var fillCount = Math.min(textFrames.length, productNames.length);

    if (textFrames.length !== productNames.length) {
        var message =
            "テキストオブジェクト: " + textFrames.length + "個\n" +
            "CSVの商品名: " + productNames.length + "件\n\n" +
            "先頭から" + fillCount + "件だけ流し込みますか？";

        if (!confirm(message)) {
            return;
        }
    }

    for (var n = 0; n < fillCount; n++) {
        textFrames[n].contents = productNames[n];
    }

    alert(
        fillCount + "件の商品名を流し込みました。\n" +
        "順序: 上から下、同じ行では左から右"
    );

    function collectTextFrames(item, output) {
        if (!item) {
            return;
        }

        if (item.typename === "TextFrame") {
            output.push(item);
            return;
        }

        if (item.typename === "GroupItem") {
            for (var j = 0; j < item.pageItems.length; j++) {
                collectTextFrames(item.pageItems[j], output);
            }
        }
    }

    // カンマ、ダブルクォート、セル内改行に対応した簡易CSVパーサー
    function parseCSV(text) {
        var result = [];
        var row = [];
        var field = "";
        var inQuotes = false;

        for (var k = 0; k < text.length; k++) {
            var ch = text.charAt(k);

            if (inQuotes) {
                if (ch === '"') {
                    if (text.charAt(k + 1) === '"') {
                        field += '"';
                        k++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field += ch;
                }
            } else {
                if (ch === '"') {
                    inQuotes = true;
                } else if (ch === ",") {
                    row.push(field);
                    field = "";
                } else if (ch === "\r") {
                    if (text.charAt(k + 1) === "\n") {
                        k++;
                    }
                    row.push(field);
                    result.push(row);
                    row = [];
                    field = "";
                } else if (ch === "\n") {
                    row.push(field);
                    result.push(row);
                    row = [];
                    field = "";
                } else {
                    field += ch;
                }
            }
        }

        if (field.length > 0 || row.length > 0) {
            row.push(field);
            result.push(row);
        }

        return result;
    }

    function trimText(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }
})();
