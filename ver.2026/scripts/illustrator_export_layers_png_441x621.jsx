#target illustrator

/*
  Illustrator: レイヤーごとにPNGを書き出す
  - 対象: トップレベルの全レイヤー
  - 出力: 441 x 621 px
  - 形式: PNG-24
  - 背景: 白（透過なし）
  - ファイル名: 01_レイヤー名.png, 02_レイヤー名.png ...
*/

(function () {
    var TARGET_WIDTH_PX = 441;
    var TARGET_HEIGHT_PX = 621;

    if (app.documents.length === 0) {
        alert("Illustratorドキュメントを開いてから実行してください。");
        return;
    }

    var doc = app.activeDocument;

    if (doc.layers.length === 0) {
        alert("書き出すレイヤーがありません。");
        return;
    }

    var outputFolder = Folder.selectDialog("PNGの出力先フォルダーを選択してください");
    if (!outputFolder) {
        return;
    }

    var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
    var artboardRect = doc.artboards[activeArtboardIndex].artboardRect;

    // artboardRect: [left, top, right, bottom]（単位はpt）
    // Illustratorでは100%書き出し時、1pt ≒ 1pxとして扱われます。
    var artboardWidth = Math.abs(artboardRect[2] - artboardRect[0]);
    var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);

    if (artboardWidth <= 0 || artboardHeight <= 0) {
        alert("アートボードのサイズを取得できませんでした。");
        return;
    }

    var horizontalScale = TARGET_WIDTH_PX / artboardWidth * 100;
    var verticalScale = TARGET_HEIGHT_PX / artboardHeight * 100;

    var sourceRatio = artboardWidth / artboardHeight;
    var targetRatio = TARGET_WIDTH_PX / TARGET_HEIGHT_PX;
    var ratioDifference = Math.abs(sourceRatio - targetRatio);

    if (ratioDifference > 0.0005) {
        var proceed = confirm(
            "現在のアートボード比率と441×621の比率が異なります。\n\n" +
            "そのまま実行すると、縦横が個別に拡大・縮小されるため、画像がわずかに変形します。\n\n" +
            "推奨: アートボードを441×621 pxに変更してから実行してください。\n\n" +
            "このまま書き出しますか？"
        );
        if (!proceed) {
            return;
        }
    }

    var originalVisibility = [];
    var exportedCount = 0;

    function zeroPad(number, digits) {
        var text = String(number);
        while (text.length < digits) {
            text = "0" + text;
        }
        return text;
    }

    function sanitizeFileName(name) {
        var result = name || "layer";

        result = result.replace(/[\\\/:\*\?"<>\|]/g, "_");
        result = result.replace(/[\r\n\t]/g, " ");
        result = result.replace(/^\s+|\s+$/g, "");
        result = result.replace(/[\. ]+$/g, "");

        if (result === "") {
            result = "layer";
        }

        if (result.length > 100) {
            result = result.substring(0, 100);
        }

        return result;
    }

    function makeWhiteColor() {
        var white = new RGBColor();
        white.red = 255;
        white.green = 255;
        white.blue = 255;
        return white;
    }

    try {
        for (var i = 0; i < doc.layers.length; i++) {
            originalVisibility[i] = doc.layers[i].visible;
        }

        var options = new ExportOptionsPNG24();
        options.antiAliasing = true;
        options.artBoardClipping = true;
        options.horizontalScale = horizontalScale;
        options.verticalScale = verticalScale;
        options.transparency = false;
        options.matte = true;
        options.matteColor = makeWhiteColor();
        options.saveAsHTML = false;

        // レイヤーパネルの上から順に書き出す
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            for (var hideIndex = 0; hideIndex < doc.layers.length; hideIndex++) {
                doc.layers[hideIndex].visible = false;
            }

            var targetLayer = doc.layers[layerIndex];
            targetLayer.visible = true;

            app.redraw();

            var fileName = sanitizeFileName(targetLayer.name);

            var outputFile = new File(outputFolder.fsName + "/" + fileName);

            doc.exportFile(outputFile, ExportType.PNG24, options);
            exportedCount++;
        }
    } catch (error) {
        alert(
            "書き出し中にエラーが発生しました。\n\n" +
            "書き出し済み: " + exportedCount + "件\n" +
            "エラー: " + error
        );
        return;
    } finally {
        for (var restoreIndex = 0; restoreIndex < doc.layers.length; restoreIndex++) {
            try {
                doc.layers[restoreIndex].visible = originalVisibility[restoreIndex];
            } catch (restoreError) {
            }
        }
        app.redraw();
    }

    alert(
        "PNGの書き出しが完了しました。\n\n" +
        "出力数: " + exportedCount + "件\n" +
        "サイズ: " + TARGET_WIDTH_PX + " × " + TARGET_HEIGHT_PX + " px\n" +
        "背景: 白（透過なし）\n" +
        "出力先: " + outputFolder.fsName
    );
})();
