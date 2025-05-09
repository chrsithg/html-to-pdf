const xl = require('excel4node');
const { config } = require('dotenv');
const express = require('express');
const puppeteer = require('puppeteer');
const fs = require('fs');
const app = express();
const bodyParser = require('body-parser');
//----------------------------------------------------
app.use(bodyParser.json({limit: '50mb'}));
app.use(bodyParser.urlencoded({limit: '50mb', extended: true}));
const busboy = require('connect-busboy');
// Loading environment variables
app.use(express.json());
app.use(busboy());
config();
//************************************************************

app.get('/api', (req, res) => {
    res.send('PDF Service Here!');
});

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
    console.log(`app is listening on port ${PORT}`);
});

//***********************************************************************
app.get('/api/upload', (req, res) => {
    const html = req.query.html
    console.log(html);
    
    //console.log(datas);
    const convert = new Convert(html);
    convert.getPdf().then(result=>{
	console.log(result);
        res.status(200).download(result);
    }).catch(err => {
        res.status(500).json(err);
    });
    //console.log(filePath);
    //res.download(filePath);
    //res.status(200).download(filepath);
    //res.send('OK');    
});

//form-data key:sample value:file
//******************************************************************************************
app.post('/api/upload', (req, res) => {
    let fstream;
    //req.pipe(req.busboy);
    if(req.busboy){
    req.busboy.on('file', (name, file, info) => {
        console.log("Uploading: " + name); 
        fstream = fs.createWriteStream(__dirname + '/upload/' + name);
        file.pipe(fstream);
        fstream.on('close', function () {
            const html = fs.readFileSync(__dirname + '/upload/' + name, 'utf-8');
	    const options = {
        	  path: './download/result.pdf',
	          margin: { top: '100px', right: '50px', bottom: '100px', left: '50px' },
	          printBackground: true,
	          format: 'A4',
	          landscape: true,
	    };
           
	    const convert = new Convert(html,options);
            convert.getPdf().then(result=>{
		res.status(200).download(result);
	    }).catch(err=>{
		res.status(500).json({'success':'failed'});
	    });
	    //res.status(200).send('ok');
        });
    });
    req.pipe(req.busboy);
    };

    //console.log(filePath);
    //res.download(filePath);
    //res.status(200).download(filepath);
    //res.send('OK');    
});

//json({base64:"xxxx"})格式
//**************************************************************************************
app.post('/api/base64', (req, res) => {
    const data = req.body.base64;
    let landscape = req.body.landscape || false;
    let html = Buffer.from(data,'base64').toString('utf8');
    const options = {
	  path: './download/result.pdf',
	  margin: { top: '100px', right: '50px', bottom: '100px', left: '50px' },
	  printBackground: true,
	  format: 'A4',
	  landscape: true,
    };
    
    const convert = new Convert(html,options);
    convert.getPdfBinary().then(result=>{
	//console.log(result);
        //res.status(200).download(result);
        res.status(200).send({"document":result});
    }).catch(err => {
        res.status(500).json({'error':err});
    });
});

//****************************************************************************************
app.get('/api/download', async (req, res) => {
	//Get HTML content from HTML file
	const html = fs.readFileSync('sample.html', 'utf-8');
	const options = {
	  path: './download/result.pdf',
	  margin: { top: '100px', right: '50px', bottom: '100px', left: '50px' },
	  printBackground: true,
          format: 'A4',
	  landscape: true,
	};

        let convert = new Convert(html,options);
        convert.getPdf().then(result => {
           res.status(200).download(result);
        })
        .catch( err => {
           res.status(500).json({'error':err}) 
        });
});
//
//******************************************************************************************
app.post('/api/xlsx', (req, res) => {
    const heading = req.body.heading;
    const jsondata = req.body.jsondata;
    //const jsonObj = JSON.parse(json);
    console.log(heading);
    console.log(jsondata);
    
    let convert = new json2xlsx(heading,jsondata);
    convert.getXlsx().then( result => {
	console.log(result);
        //res.status(200).send({"document":result});
        res.status(200).download(result);
    }).catch(err => {
        res.status(500).json({'error':err});
    });
});

//******************************************************************************************
app.post('/api/base64xlsx', (req, res) => {
    const heading = req.body.heading;
    const jsondata = req.body.jsondata;
    //const jsonObj = JSON.parse(json);
    console.log(heading);
    console.log(jsondata);

    let convert = new json2xlsx(heading,jsondata);
    convert.getXlsxBase64().then( result => {
        console.log(result);
        res.status(200).send({"document":result});
    }).catch(err => {
        res.status(500).json({'error':err});
    });
});



//***************************************************************************************
//
// function json2xlsx()
//*************************************************************************************
function json2xlsx(heading,json){
	this._heading = heading;
	this._json = json;
        this._base64String ="";

	this.getXlsx = async () => {
	   try{
		const myStyle = {
	        	alignment:{horizontal:'left', vertical:'center'},
	        	font: {size: 12}
		};
		const wb = new xl.Workbook();
		const style = wb.createStyle(myStyle);
		const ws = wb.addWorksheet('Worksheet');
                 

		//Write Column Title in Excel file
		let headingColumnIndex = 1;
		this._heading.forEach( heading => {
		   // ws.column(headingColumnIndex++).setWidth(100);
		    ws.cell(1, headingColumnIndex++).string(heading);
		});

		//Write Data in Excel file
		let rowIndex = 2;
		let data = this._json;
		await data.forEach( record => {
		    let columnIndex = 1;
		    Object.keys(record ).forEach(columnName =>{
		        ws.cell(rowIndex,columnIndex++).string(record [columnName]).style(style)
		    });
		    rowIndex++;
		});

                await wb.write('download/output.xlsx');	       
                let filePath = `${__dirname}/download/output.xlsx`;
		return filePath;
           }
	   catch(error) {
		return error
	   }
	}

      this.getXlsxBase64 = async ()=> {
         try{
                const myStyle = {
                        alignment:{horizontal:'left', vertical:'center'},
                        font: {size: 12}
                };
                const wb = new xl.Workbook();
                const style = wb.createStyle(myStyle);
                const ws = wb.addWorksheet('Worksheet');


                //Write Column Title in Excel file
                let headingColumnIndex = 1;
                this._heading.forEach( heading => {
                   // ws.column(headingColumnIndex++).setWidth(100);
                    ws.cell(1, headingColumnIndex++).string(heading);
                });

                //Write Data in Excel file
                let rowIndex = 2;
                let data = this._json;
                await data.forEach( record => {
                    let columnIndex = 1;
                    Object.keys(record ).forEach(columnName =>{
                        ws.cell(rowIndex,columnIndex++).string(record [columnName]).style(style)
                    });
                    rowIndex++;
                });

                  //** Create buffer from workbook
                await wb.writeToBuffer().then((buffer) => {
                    //const base64String = buffer.toString('base64');
                    //console.log(base64String);
                    this._base64String = buffer.toString('base64');

                  });

		return this._base64String;


                //await wb.write('download/output.xlsx');
                //let filePath = `${__dirname}/download/output.xlsx`;
		//const filedata = await fs.readFileSync(filePath);
		//const base64 = await Buffer.from(filedata).toString('base64');
		//console.log(base64);
		//return base64;
           }
           catch(error) {
                return error
           }

      }
}




//**********************************************************************************************************
// Convert()方法：
// html:
// options:
//
//***********************************************************************************************************
function Convert(html,options){
    this._html = html;
    this._options = options;

    this.getPdfBinary = async () => {
            // Create a browser instance
            const browser = await puppeteer.launch({args: ["--no-sandbox", "--disabled-setupid-sandbox", '--font-render-hinting=none']});

            // Create a new page
            const page = await browser.newPage();

            //Get HTML content from HTML file
            //const html = fs.readFileSync('sample.html', 'utf-8');
            
            await page.setContent(this._html, { waitUntil: 'domcontentloaded' });
	    //await page.addStyleTag({ content:'.trbg { background: #eeeeee }'});


	   // Add a style tag to the page with CSS to change the background color
	   await page.addStyleTag({
	     content: 'body { background-color: #ffffff; }'
	   });


            // To reflect CSS used for screens instead of print: print / screen
            await page.emulateMediaType('print');

            // 00.Downlaod the PDF
            //const pdf = await page.pdf({
            //    path: './download/result.pdf',
            //    margin: { top: '100px', right: '50px', bottom: '100px', left: '50px' },
            //    printBackground: true,
            //    format: 'A4'
            //});

            // 01.Download the PDF
            const pdf = await page.pdf(this._options);

            // Close the browser instance
            await browser.close();
            let filePath = `${__dirname}/download/result.pdf`;	    
            try {
		  const data = fs.readFileSync(filePath);
		  const base64 = Buffer.from(data).toString('base64');
		  //console.log(base64);
		  return base64;
	    } catch (err) {
		  //console.error(err);
		  return err;
	    }
    }

    this.getPdf = async () => {
            // Create a browser instance
            //const browser = await puppeteer.launch();
            const browser = await puppeteer.launch({args: ["--no-sandbox", "--disabled-setupid-sandbox", '--font-render-hinting=none']});

            // Create a new page
            const page = await browser.newPage();

            //Get HTML content from HTML file
            //const html = fs.readFileSync('sample.html', 'utf-8');
            await page.setContent(this._html, { waitUntil: 'domcontentloaded' });

           // Add a style tag to the page with CSS to change the background color
           await page.addStyleTag({
             content: 'body { background-color: #ffffff; }'
           });

           // Make watermark 浮水印
           //await page.addStyleTag({path: 'watermark.css'});


            // To reflect CSS used for screens instead of print
            await page.emulateMediaType('print');

            // Downlaod the PDF
            //const pdf = await page.pdf({
            //    path: './download/result.pdf',
            //    margin: { top: '100px', right: '50px', bottom: '100px', left: '50px' },
            //    printBackground: false,
            //    format: 'A4',
            //    landscape: true,
            //});

            // Download the PDF
            const pdf = await page.pdf(this._options);
            // Close the browser instance
            await browser.close();
            var filePath = `${__dirname}/download/result.pdf`;
            return filePath;
    }
}

