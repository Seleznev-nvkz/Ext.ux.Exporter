/**
 * @class Ext.ux.Exporter.XlsxFormatter
 * @extends Ext.ux.Exporter.Formatter
 * Specialised Format class for outputting .xlsx files
 */
Ext.define("Ext.ux.exporter.xlsxFormatter.XlsxFormatter", {
    extend: "Ext.ux.exporter.Formatter",
    uses: [
        "Ext.ux.exporter.xlsxFormatter.Worksheet"
    ],
    contentType: 'data:application/vnd.ms-excel;base64,',
    extension: "xlsx",
    _urlLibrary: 'xlsx.min.js',  // from https://github.com/Seleznev-nvkz/xlsx.js

    loadRemote: function () {
        if (typeof xlsx === "undefined") {
            Ext.Loader.loadScript({url: this._urlLibrary, onLoad: Ext.emptyFn});
        }
    },

    format: function (store, config) {
        var newSheet = new Ext.ux.exporter.xlsxFormatter.Worksheet(store, config);

        if (typeof xlsx === "undefined") {
            Ext.toast({title: 'Xlsx loading...', width: 300, align: 'tr'}).show();
        } else {
            return xlsx({
                worksheets: [newSheet.render()]
            });
        }
    }
});