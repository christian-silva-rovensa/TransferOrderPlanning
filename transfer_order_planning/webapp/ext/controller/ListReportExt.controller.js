sap.ui.define([
    "sap/m/MessageToast",
    "sap/m/MessagePopover",
    "sap/m/MessagePopoverItem",
    "rovensa/transferorderplanning/libs/moment-with-locales.min",
    "rovensa/transferorderplanning/libs/xlsx.full.min"
], function(MessageToast, MessagePopover, MessagePopoverItem) {
    'use strict';

    return {

        FileUpload: function(oEvent) {
            
            this.onOpenDialog();         

        },

        async onOpenDialog() {

			// Create dialog lazily
			this.oDialog ??= await this.loadFragment({
				name: "rovensa.transferorderplanning.ext.fragment.UploadDialog",
                controller: this
			});

			this.oDialog.open();
		},

        onCloseDialog: function(oEvent) {
            var oFileUploader = this.byId("fileUploader");
            oFileUploader.clear();
            this.oDialog.close();
        },

        handleFileChange : function(oEvent){
            this._file = oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        },

        handleFileTypeMissmatch: function(oEvent) {
            MessageToast.show("Only Excel files allowed");
        },

        handleUploadPress: function(oEvent) {

            if (!this._file) {
                return;
            }

            var oMessageManager = sap.ui.getCore().getMessageManager();

            oMessageManager.removeAllMessages();

            var that = this;
            
            var oFileUploader = this.byId("fileUploader");
			
            oFileUploader.checkFileReadable().then(async function() {
		              
                that._handleFileAsync(that.getView().getModel("custom"));

                that.oDialog.close();

			}, function(error) {
				MessageToast.show("The file cannot be read. It may have changed.");
			}).then(function() {
				oFileUploader.clear();
			});
        },

        async _handleFileAsync(oModel) {

            if (!this._file){
                return;
            }

            let that = this;
            
            const data = await this._file.arrayBuffer();
            
            const workbook = XLSX.read(data);

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
            const raw_data = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});

            //console.log(raw_data);

            if (raw_data.length < 2){ 
                return;
            }

            var aData = [];

            //slice(1) is used here to skip the header
            raw_data.slice(1).forEach(purchaseOrderItem => {

                let oData = { PurchaseOrder: purchaseOrderItem[0], 
                              PurchaseOrderItem: purchaseOrderItem[1],
                              RequestFinishDate: purchaseOrderItem[7],
                              ConfirmedLoadingDate: purchaseOrderItem[8],
                              ConfirmedLoadingTime: purchaseOrderItem[9],
                              FinalFinishDate: purchaseOrderItem[10],
                              VesselEtd: purchaseOrderItem[11],
                              MinArrivalDate: purchaseOrderItem[12],
                              MaxArrivalDate: purchaseOrderItem[13],
                              JustArrivalDate: purchaseOrderItem[14],
                              JustRequestFinishDate: purchaseOrderItem[15],
                              FinalFinishDateNotes: purchaseOrderItem[16],
                              JustFinalFinishDate: purchaseOrderItem[17],
                              ConfirmedLoadingDateNotes: purchaseOrderItem[18],
                              TransportCompany: purchaseOrderItem[19],
                              TransportPlate: purchaseOrderItem[20]
                            };

                aData.push(oData);

            });

            var aPurchaseOrdersItems = [];

            aData.forEach(entity => {

                /*
                if(entity.ExpectedAvailDate !== undefined){
                    entity.ExpectedAvailDate = moment(entity.ExpectedAvailDate).format('YYYY-MM-DDTHH:MM:SS');
                }
                
                if(entity.LoadingTime !== undefined){
                    let str = entity.LoadingTime.toString();
                    let matches = str.replace(/[^0-9:]/g, "");
                    const myArr = matches.split(":");
                    if(myArr.length === 3){
                        entity.LoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M"+myArr[2]+"S";
                    }else if(myArr.length === 2){
                        entity.LoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M00S";
                    }else{
                        entity.LoadingTime = "PT00H00M00S";
                    }
                }

                if(entity.VesselETD !== undefined){
                    entity.VesselETD = moment(entity.VesselETD).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MinDeliveryDate !== undefined){
                    entity.MinDeliveryDate = moment(entity.MinDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MaxDeliveryDate !== undefined){
                    entity.MaxDeliveryDate = moment(entity.MaxDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }
                

                var oSalesOrderItem = { SalesOrder: entity.SalesOrder, 
                                        SalesOrderItem: entity.SalesOrderItem, 
                                        ExpectedAvailDate: entity.ExpectedAvailDate,
                                        StockAvailDate: entity.StockAvailDate,
                                        ConfirmedLoadingDate: entity.ConfirmedLoadingDate,
                                        LoadingTime: entity.LoadingTime,
                                        LoadingDate: entity.LoadingDate,
                                        VesselETD: entity.VesselETD,
                                        FinalDelivery: entity.FinalDelivery,
                                        MinDeliveryDate: entity.MinDeliveryDate,
                                        MaxDeliveryDate: entity.MaxDeliveryDate,
                                        Location: entity.Location,
                                        JustExpectedAvailDate: that._addLeadingZero(entity.JustExpectedAvailDate),
                                        JustStockAvailDate: that._addLeadingZero(entity.JustStockAvailDate),
                                        JustLoadingDate: that._addLeadingZero(entity.JustLoadingDate),
                                        JustMinDeliveryDate: that._addLeadingZero(entity.JustMinDeliveryDate),
                                        JustMaxDeliveryDate: that._addLeadingZero(entity.JustMaxDeliveryDate),
                                        StockAvailDateNotes: entity.StockAvailDateNotes,
                                        LoadingDateNotes: entity.LoadingDateNotes
                                     };
                
                aSalesOrderItems.push(oSalesOrderItem);

                */

                if(entity.RequestFinishDate !== undefined){
                    entity.RequestFinishDate = moment(entity.RequestFinishDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.ConfirmedLoadingDate !== undefined){
                    entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.ConfirmedLoadingTime !== undefined){
                    let str = entity.ConfirmedLoadingTime.toString();
                    let matches = str.replace(/[^0-9:]/g, "");
                    const myArr = matches.split(":");
                    if(myArr.length === 3){
                        entity.ConfirmedLoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M"+myArr[2]+"S";
                    }else if(myArr.length === 2){
                        entity.ConfirmedLoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M00S";
                    }else{
                        entity.ConfirmedLoadingTime = "PT00H00M00S";
                    }
                }

                if(entity.FinalFinishDate !== undefined){
                    entity.FinalFinishDate = moment(entity.FinalFinishDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.VesselEtd !== undefined){
                    entity.VesselEtd = moment(entity.VesselEtd).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MinArrivalDate !== undefined){
                    entity.MinArrivalDate = moment(entity.MinArrivalDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MaxArrivalDate !== undefined){
                    entity.MaxArrivalDate = moment(entity.MaxArrivalDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                var oPurchaseOrderItem = { PurchaseOrder: entity.PurchaseOrder, 
                                           PurchaseOrderItem: entity.PurchaseOrderItem,
                                           RequestFinishDate: entity.RequestFinishDate,
                                           ConfirmedLoadingDate: entity.ConfirmedLoadingDate,
                                           ConfirmedLoadingTime: entity.ConfirmedLoadingTime,
                                           FinalFinishDate: entity.FinalFinishDate,
                                           VesselEtd: entity.VesselEtd,
                                           MinArrivalDate: entity.MinArrivalDate,
                                           MaxArrivalDate: entity.MaxArrivalDate,
                                           JustArrivalDate: entity.JustArrivalDate,
                                           JustRequestFinishDate: entity.JustRequestFinishDate,
                                           FinalFinishDateNotes: entity.FinalFinishDateNotes,
                                           JustFinalFinishDate: entity.JustFinalFinishDate,
                                           ConfirmedLoadingDateNotes: entity.ConfirmedLoadingDateNotes,
                                           TransportCompany: entity.TransportCompany,
                                           TransportPlate: entity.TransportPlate
                                        };
                
                aPurchaseOrdersItems.push(oPurchaseOrderItem);

            });

            if( aPurchaseOrdersItems.length > 0){
                MessageToast.show("Planning is being updated");
            }

            aPurchaseOrdersItems.forEach(entity => {
            
                // Bypass ETag optimistic locking when updating entities via file upload
                oModel.setHeaders({"If-Match": "*"});

                oModel.update("/Planning(PurchaseOrder='"+entity.PurchaseOrder+"',PurchaseOrderItem='"+entity.PurchaseOrderItem+"')",
                    entity, {
                    success: function(response){
                        
                        if(!response){
                            let sMessage = "Purchase Order: "+ entity.PurchaseOrder + ", Item: " + entity.PurchaseOrderItem + " updated successfully.";
                            that.addMessage(sMessage, "S");
                        }

                    }, 
                    error: function(err){
                        
                        /*
                        if(err.responseText){

                            let oError = JSON.parse(err.responseText);

                            if (oError.error.message.value) {
                                // Add a message only if the error message does not contain the Sales Ordem information  
                                if (!oError.error.message.value.includes(entity.SalesOrder)) {
                                    let sMessage = "Sales Order: "+ entity.SalesOrder + ", Sales Order Item: " + entity.SalesOrderItem + " - " + oError.error.message.value;
                                    that.addMessage(sMessage, "E");
                                }
                            }
                            
                        }
                        */

                    }
                });

            });

            that.DisplayMessages();

        },

        DisplayMessages: function(oEvent) {

            var oMP = new MessagePopover({
                items: {
                    path:"message>/",
                    template: new MessagePopoverItem({ description: "{message>description}", type: "{message>type}", title: "{message>message}"})
                }
            });

            var oMessageManager = sap.ui.getCore().getMessageManager();
            
            oMP.setModel(oMessageManager.getMessageModel(),"message");

            var oDisplayMessagesButton = this.byId("DisplayMessagesButton");

            oMP.openBy(oDisplayMessagesButton);
        },
        
        DownloadTemplate: function(oEvent) {

            let aContexts = this.extensionAPI.getSelectedContexts();

            var oData = [];

            if (aContexts.length > 0) {

                aContexts.forEach(element => {
                    
                    oData.push({ PurchaseOrder: element.getProperty("PurchaseOrder"), 
                                 PurchaseOrderItem: element.getProperty("PurchaseOrderItem"), 
                                 Material: element.getProperty("Material"),
                                 MaterialName: element.getProperty("MaterialName"),
                                 OrderQuantity: element.getProperty("OrderQuantity"),
                                 PurchaseOrderQuantityUnit: element.getProperty("PurchaseOrderQuantityUnit"),
                                 FirstDeliveryDate: element.getProperty("FirstDeliveryDate"),
                                 RequestFinishDate: element.getProperty("RequestFinishDate"),
                                 ConfirmedLoadingDate: element.getProperty("ConfirmedLoadingDate"),
                                 ConfirmedLoadingTime: element.getProperty("ConfirmedLoadingTime"),
                                 FinalFinishDate: element.getProperty("FinalFinishDate"),
                                 VesselEtd: element.getProperty("VesselETD"),
                                 MinArrivalDate: element.getProperty("MinArrivalDate"),
                                 MaxArrivalDate: element.getProperty("MaxArrivalDate"),
                                 ArrivalDateJ: element.getProperty("ArrivalDateJ"),
                                 RequestFinishDateJ: element.getProperty("RequestFinishDateJ"),
                                 FinalFinishDateNotes: element.getProperty("FinalFinishDateNotes"),
                                 FinalFinishDateJ: element.getProperty("FinalFinishDateJ"),
                                 ConfirmedLoadingDateNotes: element.getProperty("ConfirmedLoadingDateNotes"),
                                 TransportCompany: element.getProperty("TransportCompany"),
                                 TransportPlate: element.getProperty("TransportPlate")
                            });

                });

            }else{
                
                //Sample data
                oData.push({ PurchaseOrder: "9999999999", 
                             PurchaseOrderItem: "10", 
                             Material: "",
                             MaterialName: "",
                             OrderQuantity: "0",
                             PurchaseOrderQuantityUnit: "",
                             FirstDeliveryDate: "31/12/9999",
                             RequestFinishDate: "31/12/9999",
                             ConfirmedLoadingDate: "31/12/9999",
                             ConfirmedLoadingTime: "23:59:59",
                             FinalFinishDate: "31/12/9999",
                             VesselEtd: "31/12/9999",
                             MinArrivalDate: "31/12/9999",
                             MaxArrivalDate: "31/12/9999",
                             JustArrivalDate: "",
                             JustRequestFinishDate: "",
                             FinalFinishDateNotes: "",
                             JustFinalFinishDate: "",
                             ConfirmedLoadingDateNotes: "",
                             TransportCompany: "",
                             TransportPlate: ""
                        });
            }

            const worksheet = XLSX.utils.json_to_sheet(oData);

            const workbook = XLSX.utils.book_new();
            
            XLSX.utils.book_append_sheet(workbook, worksheet, "Planning");  
            
            const columnLetter = 'E'; "Order Quantity"
            const range = XLSX.utils.decode_range(worksheet['!ref']);

            // Iterate through each cell in the column order quantity and set the type to 'n' (number)
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                const cellAddress = `${columnLetter}${row + 1}`;
                if (worksheet[cellAddress]) {
                    worksheet[cellAddress].t = 'n'; // Set cell type to number
                }
            }
            
            XLSX.writeFile(workbook, "PlanningTemplate.xlsx", { compression: true });

        },

        _addLeadingZero: function(sValue) {

            if (sValue === undefined) {
                return;
            }

            let sValueNew = sValue; 

            const regex = new RegExp('[0-9]');

            try {
              
                if(sValueNew.length === 1 && regex.test(sValueNew) === true) {
                    sValueNew = '0'+sValueNew;
                }
              
            } catch (error) {
                console.log(error);        
            }

            return sValueNew;

        },

        _getMessageType: function(sType) {

            var sMessageType = sap.ui.core.MessageType.None;

            switch (sType) {
             
                case "S":
                    sMessageType = sap.ui.core.MessageType.Success;
                    break;
                
                case "E":
                    sMessageType = sap.ui.core.MessageType.Error;
                    break;
                
                case "W":
                    sMessageType = sap.ui.core.MessageType.Warning;
                    break;
                
                case "I":
                    sMessageType = sap.ui.core.MessageType.Information;
                    break;

            }

            return sMessageType;

        },

        _setCustomMessage: function(sMessage, sType) {

            var oMessage = {
                message: sMessage,
                type: this._getMessageType(sType)
            };

            this.extensionAPI.setCustomMessage(oMessage);

        },

        addMessage: function(sMessage, sType) {

            var oMessageManager = sap.ui.getCore().getMessageManager();

            var oMessage = new sap.ui.core.message.Message({
                message: sMessage,
                persistent: true, // create message as transition message
                type: this._getMessageType(sType)
            });

            oMessageManager.addMessages(oMessage);

        }
    };
});