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
            
            const workbook = XLSX.read(data, {cellDates: true, dateNF:"dd/mm/yyyy"});

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
                              JustRequestFinishDate: purchaseOrderItem[8],
                              FinalFinishDate: purchaseOrderItem[9],
                              JustFinalFinishDate: purchaseOrderItem[10],
                              FinalFinishDateNotes: purchaseOrderItem[11],
                              ConfirmedLoadingDate: purchaseOrderItem[12],
                              ConfirmedLoadingTime: purchaseOrderItem[13],
                              ConfirmedLoadingDateNotes: purchaseOrderItem[14],
                              ExpectedLoadingDate: purchaseOrderItem[15],
                              JustExpectedLoadingDate: purchaseOrderItem[16],
                              FinalDelivery: purchaseOrderItem[17],
                              MinArrivalDate: purchaseOrderItem[18],
                              MaxArrivalDate: purchaseOrderItem[19],
                              JustArrivalDate: purchaseOrderItem[20],
                              TransportCompany: purchaseOrderItem[21],
                              TransportPlate: purchaseOrderItem[22],
                              Location: purchaseOrderItem[23]
                            };

                if(oData.PurchaseOrder !== undefined) {
                    aData.push(oData);
                }

            });

            var aPurchaseOrdersItems = [];

            aData.forEach(entity => {

                let arrAcceptedDateFormats = ["DD/MM/YYYY","MM/DD/YYYY","DD.MM.YYYY","MM.DD.YYYY","YYYY-MM-DD"];

                if(entity.RequestFinishDate !== undefined){
                    //entity.RequestFinishDate = moment(entity.RequestFinishDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.RequestFinishDate, arrAcceptedDateFormats).isValid()){
                        entity.RequestFinishDate = moment(entity.RequestFinishDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.RequestFinishDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.RequestFinishDate = null;
                        }else{
                            entity.RequestFinishDate = undefined;
                        }
                    }
                }

                if(entity.ConfirmedLoadingDate !== undefined){
                    //entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.ConfirmedLoadingDate, arrAcceptedDateFormats).isValid()){
                        entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.ConfirmedLoadingDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.ConfirmedLoadingDate = null;
                        }else{
                            entity.ConfirmedLoadingDate = undefined;
                        }
                    }
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

                if(entity.RequestFinishDate !== undefined){
                    //entity.RequestFinishDate = moment(entity.RequestFinishDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.RequestFinishDate, arrAcceptedDateFormats).isValid()){
                        entity.RequestFinishDate = moment(entity.RequestFinishDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.RequestFinishDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.RequestFinishDate = null;
                        }else{
                            entity.RequestFinishDate = undefined;
                        }
                    }
                }

                if(entity.FinalDelivery !== undefined){
                    //entity.FinalDelivery = moment(entity.FinalDelivery).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.FinalDelivery, arrAcceptedDateFormats).isValid()){
                        entity.FinalDelivery = moment(entity.FinalDelivery, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.FinalDelivery == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.FinalDelivery = null;
                        }else{
                            entity.FinalDelivery = undefined;
                        }
                    }
                }

                if(entity.FinalFinishDate !== undefined){
                    //entity.FinalFinishDate = moment(entity.FinalFinishDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.FinalFinishDate, arrAcceptedDateFormats).isValid()){
                        entity.FinalFinishDate = moment(entity.FinalFinishDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.FinalFinishDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.FinalFinishDate = null;
                        }else{
                            entity.FinalFinishDate = undefined;
                        }
                    }
                }

                if(entity.ConfirmedLoadingDate !== undefined){
                    //entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.ConfirmedLoadingDate, arrAcceptedDateFormats).isValid()){
                        entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.ConfirmedLoadingDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.ConfirmedLoadingDate = null;
                        }else{
                            entity.ConfirmedLoadingDate = undefined;
                        }
                    }
                }
                
                if(entity.ExpectedLoadingDate !== undefined){
                    //entity.ExpectedLoadingDate = moment(entity.ExpectedLoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.ExpectedLoadingDate, arrAcceptedDateFormats).isValid()){
                        entity.ExpectedLoadingDate = moment(entity.ExpectedLoadingDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.ExpectedLoadingDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.ExpectedLoadingDate = null;
                        }else{
                            entity.ExpectedLoadingDate = undefined;
                        }
                    }
                }

                if(entity.MinArrivalDate !== undefined){
                    //entity.MinArrivalDate = moment(entity.MinArrivalDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.MinArrivalDate, arrAcceptedDateFormats).isValid()){
                        entity.MinArrivalDate = moment(entity.MinArrivalDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.MinArrivalDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.MinArrivalDate = null;
                        }else{
                            entity.MinArrivalDate = undefined;
                        }
                    }
                }

                if(entity.MaxArrivalDate !== undefined){
                    //entity.MaxArrivalDate = moment(entity.MaxArrivalDate).format('YYYY-MM-DDTHH:MM:SS');
                    if(moment(entity.MaxArrivalDate, arrAcceptedDateFormats).isValid()){
                        entity.MaxArrivalDate = moment(entity.MaxArrivalDate, arrAcceptedDateFormats).format('YYYY-MM-DDTHH:MM:SS');
                    }else{
                        if(entity.MaxArrivalDate == '00/00/0000'){
                            // clears the field value on the backend 
                            entity.MaxArrivalDate = null;
                        }else{
                            entity.MaxArrivalDate = undefined;
                        }
                    }
                }

                var oPurchaseOrderItem = { PurchaseOrder: entity.PurchaseOrder, 
                                           PurchaseOrderItem: entity.PurchaseOrderItem,
                                           RequestFinishDate: entity.RequestFinishDate,
                                           RequestFinishDateJ: that._addLeadingZero(entity.JustRequestFinishDate),
                                           FinalFinishDate: entity.FinalFinishDate,
                                           FinalFinishDateJ: that._addLeadingZero(entity.JustFinalFinishDate),
                                           FinalFinishDateNotes: entity.FinalFinishDateNotes,
                                           ConfirmedLoadingDate: entity.ConfirmedLoadingDate,
                                           ConfirmedLoadingTime: entity.ConfirmedLoadingTime,
                                           ConfirmedLoadingDateNotes: entity.ConfirmedLoadingDateNotes,
                                           ExpLoadingDate: entity.ExpectedLoadingDate,     
                                           ExpLoadingDateJ: that._addLeadingZero(entity.JustExpectedLoadingDate),                                      
                                           FinalDelivery: entity.FinalDelivery,
                                           MinArrivalDate: entity.MinArrivalDate,
                                           MaxArrivalDate: entity.MaxArrivalDate,
                                           ArrivalDateJ: that._addLeadingZero(entity.JustArrivalDate),
                                           TransportCompany: entity.TransportCompany,
                                           TransportPlate: entity.TransportPlate,
                                           Location: entity.Location
                                        };
                
                aPurchaseOrdersItems.push(oPurchaseOrderItem);

            });

            if( aPurchaseOrdersItems.length > 0){
                MessageToast.show("Planning is being updated");
            }else{
                MessageToast.show("No valid entries found");
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

            var aData = [];

            if (aContexts.length > 0) {

                aContexts.forEach(element => {

                    let oConfirmedLoadingTime = element.getProperty("ConfirmedLoadingTime");

                    if (oConfirmedLoadingTime !== undefined) {
                        if (oConfirmedLoadingTime.ms !== undefined) {
                            oConfirmedLoadingTime = this._msToTime(oConfirmedLoadingTime.ms);
                        }
                    }

                    aData.push({ PurchaseOrder: element.getProperty("PurchaseOrder"), 
                                 PurchaseOrderItem: element.getProperty("PurchaseOrderItem"), 
                                 Material: element.getProperty("Material"),
                                 MaterialName: element.getProperty("MaterialName"),
                                 OrderQuantity: element.getProperty("OrderQuantity"),
                                 PurchaseOrderQuantityUnit: element.getProperty("PurchaseOrderQuantityUnit"),
                                 PlantName: element.getProperty("PlantName"),
                                 RequestFinishDate: element.getProperty("RequestFinishDate"),
                                 JustRequestFinishDate: element.getProperty("RequestFinishDateJ"),
                                 FinalFinishDate: element.getProperty("FinalFinishDate"),
                                 JustFinalFinishDate: element.getProperty("FinalFinishDateJ"),
                                 FinalFinishDateNotes: element.getProperty("FinalFinishDateNotes"),
                                 ConfirmedLoadingDate: element.getProperty("ConfirmedLoadingDate"),
                                 ConfirmedLoadingTime: oConfirmedLoadingTime,
                                 ConfirmedLoadingDateNotes: element.getProperty("ConfirmedLoadingDateNotes"),
                                 ExpectedLoadingDate: element.getProperty("ExpLoadingDate"),
                                 JustExpectedLoadingDate: element.getProperty("ExpLoadingDateJ"),
                                 FinalDelivery: element.getProperty("FinalDelivery"),
                                 MinArrivalDate: element.getProperty("MinArrivalDate"),
                                 MaxArrivalDate: element.getProperty("MaxArrivalDate"),
                                 JustMaxArrivalDate: element.getProperty("ArrivalDateJ"),
                                 TransportCompany: element.getProperty("TransportCompany"),
                                 TransportPlate: element.getProperty("TransportPlate"),
                                 Location: element.getProperty("Location")
                            });

                });

            }else{
                
                //Sample data
                aData.push({ PurchaseOrder: "9999999999", 
                             PurchaseOrderItem: "10", 
                             Material: "",
                             MaterialName: "",
                             OrderQuantity: "0",
                             PurchaseOrderQuantityUnit: "",
                             PlantName: "",
                             RequestFinishDate: "31/12/9999",
                             JustRequestFinishDate: "",
                             FinalFinishDate: "31/12/9999",
                             JustFinalFinishDate: "",
                             FinalFinishDateNotes: "",
                             ConfirmedLoadingDate: "31/12/9999",
                             ConfirmedLoadingTime: "23:59:59",
                             ConfirmedLoadingDateNotes: "",
                             ExpectedLoadingDate: "31/12/9999",
                             JustExpectedLoadingDate: "",
                             FinalDelivery: "31/12/9999",
                             MinArrivalDate: "31/12/9999",
                             MaxArrivalDate: "31/12/9999",
                             JustMaxArrivalDate: "",
                             TransportCompany: "",
                             TransportPlate: "",
                             Location: ""
                        });
      
            }

            const worksheet = XLSX.utils.json_to_sheet(aData);

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

        _msToTime(ms) {
            const totalSeconds = Math.floor(ms / 1000);
            const hours = Math.floor(totalSeconds / 3600);
            const minutes = Math.floor((totalSeconds % 3600) / 60);
            const seconds = totalSeconds % 60;

            return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
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