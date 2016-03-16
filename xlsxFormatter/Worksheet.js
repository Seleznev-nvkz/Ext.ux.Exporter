/**
 * @class Ext.ux.exporter.xlsxFormatter.Worksheet
 * @extends Object
 * Represents an Excel worksheet
 * @cfg {Ext.data.Store} store The store to use (required)
 */
Ext.define("Ext.ux.exporter.xlsxFormatter.Worksheet", {

    constructor: function (store, config) {
        config = config || {};

        this.store = store;

        Ext.applyIf(config, {
            columns: store.fields == undefined ? {} : store.fields.items,
            defaultBorders: {left: 'e4e4e4', right: 'e4e4e4', top: 'e4e4e4', bottom: 'e4e4e4'}
        });

        Ext.apply(this, config);

        Ext.ux.exporter.xlsxFormatter.Worksheet.superclass.constructor.apply(this, arguments);
    },

    render: function () {
        var data = [
            [{
                colSpan: this.columns.length,
                value: 'Result',
                vAlign: 'center',
                hAlign: 'center',
                fontSize: 15,
                fontName: 'arial',
                bold: 1
            }],
            this.buildHeader()
        ];
        return {
            data: data.concat(this.buildRows()),
            name: 'Result'
        };
    },

    buildRows: function () {
        var rows = [],
            odd = false;

        this.store.each(function (record, index) {
            rows.push(this.buildRow(record, index, odd));
            odd = !odd;
        }, this);

        return rows;
    },

    buildHeader: function () {
        var cells = [],
            cell = {
                fontSize: 10,
                bold: 1,
                WrapText: 1,
                fontName: 'arial',
                vAlign: 'center',
                hAlign: 'center',
                backgroundColor: 'A3C9F1',
                autoWidth: true,
                borders: this.defaultBorders
            };

        Ext.each(this.columns, function (col) {
            var current_cell = Ext.clone(cell),
                title;

            if (col.text != undefined) {
                title = col.text;
            } else if (col.name) {
                title = col.name.replace(/_/g, " ");
                title = Ext.String.capitalize(title);
            }
            current_cell['value'] = title;
            cells.push(current_cell);
        }, this);

        return cells;
    },

    buildRow: function (record, index, odd) {
        var cells = [],
            iCol = 0;

        Ext.each(this.columns, function (col) {
            var name = col.name || col.dataIndex,
                value;

            if (!Ext.isEmpty(name)) {
                if (Ext.isFunction(col.renderer)) {
                    value = col.renderer(record.get(name), {column: col}, record, index, iCol++, this.store);
                }
                if (Ext.isEmpty(value)) {
                    value = record.get(name);
                }
                cells.push({
                    value: value,
                    fontSize: 10,
                    formatCode: "@",
                    fontName: 'arial',
                    borders: this.defaultBorders,
                    width: 15,
                    backgroundColor: odd ? 'CCCCFF' : 'CCFFFF'
                });
            }
        }, this);

        return cells;
    }
});