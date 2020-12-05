//BPIT -> request

//*npm install request
let request=require("request");

/*npm install cheerio*/
let ch=require("cheerio");

/*file system (to use it we need to require it)*/ 
let fs=require("fs");
let path=require("path");
let xlsx=require("xlsx");

request("http://www.bpitindia.com/contact.html",getMatchurl)
function getMatchurl(err,resp,html)
{
    let STool=ch.load(html);
    let allProfile=STool("a");
    for(let i=0;i<allProfile.length/2;i++)
    {
        if(STool(allProfile[i]).text()=="Faculty Profile")
        {
            let url=STool(allProfile[i]).attr("href");
            let fUrl="http://www.bpitindia.com/"+url;
            findDataOfBranch(fUrl);
        }
    }
}

function findDataOfBranch(url)
{
    request(url,UrlKaAns);

    function UrlKaAns(err,res,html)
    {
        let STool=ch.load(html);

        let AllNames=STool("div.col-md-4");
        let AllDetails=STool("div.col-md-8");
        
        let branch=STool("h4.course-title").text();
        let b=branch.split("Departments");
        branch=b[0].trim();
        for(let i=0;i<AllNames.length;i++)
        {
            let name=STool(AllNames[i]).find("a.d_inline.fw_600").text();
            let table=STool(AllDetails[i]).find("table.table.table-bordered.table").find("tbody tr");
            let rCol=STool(table).find("td");

            let qual=STool(rCol[0]).text().trim();
            let email=STool(rCol[2]).text().trim();
            let exp=STool(rCol[4]).text().trim();
            let research=STool(rCol[6]).text().trim();
            let public=STool(rCol[8]).text().trim();
            let inter=STool(rCol[10]).text().trim();
            console.log(`Name=${name} Qualification=${qual} Email=${email} Experience=${exp} Research=${research} Publication=${public} International Publication=${inter}`);
            console.log("*************************************************************************************************");
            process(branch,name,qual,email,exp,research,public,inter);
        }
         
    }

}

function process(branch,name,qual,email,exp,research,public,inter)
{
    let dirPath=branch;
    let pStats={
        Branch: branch,
        Name: name,
        Qualification: qual,
        Email: email,
        Experience: exp,
        Research: research,
        Publication: public,
        International_Publication: inter,
    }
    
    //name=name.split("").join("");
    if(fs.existsSync(dirPath))
    {
        // file check
        // console.log("folder exist");
    }
    else
    {
        //create folder
        fs.mkdirSync(dirPath);
        //console.log("Team name folder created.");
    }

    //create file
    let FilePath=path.join(dirPath,name+".xlsx");
    let pData=[];

    if (fs.existsSync(FilePath))
    {
        pData=excelReader(FilePath,name);
        pData.push(pStats);
    }
    else
    {
        //file is created
        console.log("File",FilePath,"created");
        pData=[pStats];
    }
    excelWriter(FilePath,pData,name);
    
    

    /*Read Excel file*/
    function excelReader(filePath,name)
    {
        if(!fs.existsSync(filePath))
        {
            return null;
        }
        else
        {
            //workbook -> excel
            let wt=xlsx.readFile(filePath);

            //get data from workbook
            let excelData=wt.Sheets[name];

            //convert excel format to JSON -> array of obj
            let ans=xlsx.utils.sheet_to_json(excelData);

            //console.log(ans);
            return ans;
        }
        
    }

    /*Write Excel file*/
    function excelWriter(filePath,json,name)
    {
        name=name.split("").join("");
        //console.log(xlsx.readFile(filePath));
        let newWB=xlsx.utils.book_new();
        //console.log(json);
        let newWS=xlsx.utils.json_to_sheet(json);
        //workbook name as a parameter
        xlsx.utils.book_append_sheet(newWB,newWS,name);
        //if file is not there then it is created. Else it is replaced.
        xlsx.writeFile(newWB,filePath);
    }
}