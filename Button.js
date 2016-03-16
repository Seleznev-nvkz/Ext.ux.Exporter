Ext.define('Ext.ux.exporter.Button', {
    extend: "Ext.Button",
    alias: "widget.exporterbutton",
    uses: [
        "Ext.ux.exporter.Base64"
    ],
    config: {
        formatter: 'xlsx'
    },

    constructor: function (config) {
        config = config || {};

        Ext.ux.exporter.Button.superclass.constructor.call(this, config);

        var me = this;
        this.on("afterrender", function () {
            me.setComponent(me.store || me.component || me.up("gridpanel") || me.up("treepanel"), config);
        });
    },

    initComponent: function () {
        var me = this,
            formatterName = me.getFormatter(),
            formatter = Ext.ux.exporter.Exporter.getFormatterByName(formatterName);

        if (formatter.loadRemote) {
            formatter.loadRemote();
        }
        me.setText("Download " + formatterName);
        me.callParent(arguments);
    },

    setComponent: function (component) {
        this.component = component;
        this.store = !component.is ? component : component.getStore();
    },

    handler: function () {
        var me = this,
            data,
            location,
            fileName,
            el;

        data = Ext.ux.exporter.Exporter.exportAny(me.component, me.getFormatter());
        location = me.getMimeType() + ',' + data;
        fileName = "Export" + "-" + Ext.Date.format(new Date(), 'Y-m-d H:i') + '.' + Ext.ux.exporter.Exporter.getFormatterByName(me.getFormatter()).extension;

        el = Ext.DomHelper.append(Ext.getBody(), {
            tag: "a",
            download: fileName,
            href: location
        });

        el.click();

        Ext.fly(el).destroy();
    },

    getMimeType: function () {
        return this.getFormatter() === 'csv' ? "data:application/csv;base64" : "data:application/vnd.ms-excel;base64";
    }
});