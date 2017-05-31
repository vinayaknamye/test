let XLSX = require('xlsx');
let extract = require('pdf-text-extract');
let fs = require('fs');
let jinqJs = require('jinq');
let json2csv = require("json2csv");
let pdf_path = "./input pdfs/";
let order_ids_in_pdfs = [];
let excel_path = "./input excel/";
let tasks = [];
let excel_data = [];
let order_id_column_in_excel = "Order ID/FSN";
let columns = [
	{
		field: 'Order ID/FSN',
		text: 'Order Id'
	},
	{
		field: 'Dispatch Date',
		text: 'Dispatch Date'
	},
	{
		field: 'Seller SKU',
		text: 'Seller SKU'
	},
	{
		field: 'Order Status',
		text: 'Order Status'
	},
	{
		field: 'Order Item Value (Rs.)',
		text: 'Order Item Value'
	},
	{
		field: 'Settlement Value (Rs.) ',
		text: 'Settlement Value'
	}
];

let extractPdfText = (file_path) => {
	return new Promise((resolve, reject) => {
		extract(file_path, { splitPages: false }, function (err, text) {
		  if (err) {
		    console.dir(err)
		    return;
		  } else if(text && text[0]) {
		  	let etexts = text[0].match(/\sod[\d]+\s+/gi);
		  	if(etexts) {
			  	etexts.forEach(t => {
			  		let id = t.trim();
			  		if(order_ids_in_pdfs.indexOf(id) === -1){
			  			order_ids_in_pdfs.push(id);
			  		}
			  	});
		  	}
		  	resolve();
		  }
		});
	});
}

fs.readdirSync(pdf_path).forEach(file => {
	if(/\.pdf$/i.test(file)) {
		tasks.push(extractPdfText(pdf_path + file));
	}
});
Promise.all(tasks)
	.then((suc) => {

		fs.readdirSync(excel_path).forEach(file => {
			if(/\.xlsx$/i.test(file) || /\.xlx$/i.test(file)) {
				let workbook = XLSX.readFile(excel_path + file);
				let first_sheet_name = workbook.SheetNames[0];
				let sheet_data = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name]);
				Array.prototype.push.apply(excel_data, sheet_data);
			}
		});

		let list_of_order_ids_in_pdf_not_in_excel = new jinqJs()
								                .from(order_ids_in_pdfs)
								                .not()
          										.in(excel_data.map((d) => {
          											return d[order_id_column_in_excel];
          										}))
								                .select();

        let list_of_orders_in_pdf_and_excel = new jinqJs()
								                .from(excel_data)
								                .in(order_ids_in_pdfs, order_id_column_in_excel)
								                .select(columns);

        let result;
        if(order_ids_in_pdfs.length) {
        	result = json2csv({
        		data: order_ids_in_pdfs.map((o) => {
        			return { "Order Id" : o};
        		})
        	});
        	try {
	        	fs.unlinkSync("Order Ids in Pdfs.csv");
	        } catch(e) {

	        }

			fs.writeFileSync("Order Ids in Pdfs.csv", result);
        } else {
        	console.log("No order ids in pdfs");
        }
        if(list_of_order_ids_in_pdf_not_in_excel.length) {
        	result = json2csv({
        		data: list_of_order_ids_in_pdf_not_in_excel.map((o) => {
        			return { "Order Id" : o};
        		})
        	});
			try {
	        	fs.unlinkSync("Order Ids present in Pdfs but not in excel.csv");
	        } catch(e) {

	        }
        	fs.writeFileSync("Order Ids present in Pdfs but not in excel.csv", result);
        }  else {
        	console.log("No Order Ids present in Pdfs but not present in excel");
        }
        if(list_of_orders_in_pdf_and_excel.length) {
        	result = json2csv({
        		data: list_of_orders_in_pdf_and_excel
        	});
        	try {
	        	fs.unlinkSync("Order Ids present both in Pdfs and excel.csv");
	        } catch(e) {

	        }
        	fs.writeFileSync("Order Ids present both in Pdfs and excel.csv", result);
        } else {
        	console.log("No Order Ids present both in Pdfs and excel");
        }
	})
	.catch((err) => {
		console.log(err);
	});


