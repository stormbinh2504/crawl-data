const express = require('express');
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const axios = require('axios');
const filename = 'dataCompany.xlsx';
const checkFileExists = (filename) => {
    return fs.existsSync(filename);
};

const createExcelFile = (header) => {
    // Tạo sheet từ dữ liệu và header
    const sheet = XLSX.utils.aoa_to_sheet([header]);

    // Tạo workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'Sheet1');

    // Lưu workbook vào file Excel
    const filePath = path.join(__dirname, filename);
    XLSX.writeFile(wb, filePath);

    console.log(`Excel file created at ${filePath}`);
};

const fetchData = async (page) => {
    const requestBody = {
        query: '',
        page: page,
        facetFilters: [],
        optionalFilters: []
    };
    try {
        const response = await axios.post(
            "https://xd0u5m6y4r-3.algolianet.com/1/indexes/event-edition-eve-20c5c4b0-ff85-4f61-975b-8a627a114533_en-gb/query?x-algolia-agent=Algolia%20for%20vanilla%20JavaScript%203.27.1&x-algolia-application-id=XD0U5M6Y4R&x-algolia-api-key=d5cd7d4ec26134ff4a34d736a7f9ad47",
            JSON.stringify(requestBody),
            {
                headers: {
                    accept: "application/json",
                    "accept-language": "vi,en;q=0.9,en-US;q=0.8",
                    "cache-control": "no-cache",
                    "content-type": "application/x-www-form-urlencoded",
                    pragma: "no-cache",
                    "sec-ch-ua": "\"Chromium\";v=\"122\", \"Not(A:Brand\";v=\"24\", \"Google Chrome\";v=\"122\"",
                    "sec-ch-ua-mobile": "?0",
                    "sec-ch-ua-platform": "\"macOS\"",
                    "sec-fetch-dest": "empty",
                    "sec-fetch-mode": "cors",
                    "sec-fetch-site": "cross-site",
                    Referer: "https://www.wsew.jp/",
                    "Referrer-Policy": "strict-origin-when-cross-origin"
                }
            }
        );
        let result = response.data.hits
        for (let i = 0; i < result.length; i++) {
            let item = result[i]
            await fetchAddressId(item.organisationGuid, item.eventEditionExternalId, item, i)
        }
        console.log("load success", response.data.hits.length);
        console.log("load page", response.data.page);
        // Xử lý dữ liệu ở đây nếu cần
    } catch (error) {
        console.error("Error fetching data:", JSON.parse(error));
    }
};

const fetchAddressId = async (organisationGuid, eventEditionExternalId, item, stt) => {
    const query = `
     {
        exhibitingOrganisation(eventEditionId:"${eventEditionExternalId}", organisationId:"${organisationGuid}"){
            id  
            multilingual {
             addressLine1
             addressLine2
          }
          filterCategories{      
          responses{
            multilingual{
              name,
              locale
            }
          },
        }    
      }
    }
    `;
    try {
        const response = await axios.post(
            "https://api.reedexpo.com/graphql/",
            {
                query: query // Truy vấn GraphQL được truyền dưới dạng đối tượng JavaScript
            },
            {
                headers: {
                    "accept": "application/json",
                    "accept-language": "vi,en;q=0.9,en-US;q=0.8",
                    "cache-control": "no-cache",
                    "content-type": "application/json",
                    "pragma": "no-cache",
                    "sec-ch-ua": "\"Chromium\";v=\"122\", \"Not(A:Brand\";v=\"24\", \"Google Chrome\";v=\"122\"",
                    "sec-ch-ua-mobile": "?0",
                    "sec-ch-ua-platform": "\"macOS\"",
                    "sec-fetch-dest": "empty",
                    "sec-fetch-mode": "cors",
                    "sec-fetch-site": "cross-site",
                    "x-clientid": "uhQVcmxLwXAjVtVpTvoerERiZSsNz0om",
                    "x-correlationid": "3f31efb0-fdd6-4285-8eef-759307360546",
                    "Referer": "https://www.wsew.jp/",
                    "Referrer-Policy": "strict-origin-when-cross-origin"
                }
            }
        );
        let product = ''
        let target = ''
        let exhibitingOrganisation = response.data.data.exhibitingOrganisation
        let address = exhibitingOrganisation.multilingual[0].addressLine1
        if (exhibitingOrganisation.filterCategories[0]) {
            if (exhibitingOrganisation.filterCategories[0].responses && exhibitingOrganisation.filterCategories[0].responses.length > 0) {
                let result = exhibitingOrganisation.filterCategories[0].responses
                for (let i = 0; i < result.length; i++) {
                    let item = result[i].multilingual
                    product += item[1].name;
                    if (i < result.length - 1) {
                        product += '||';
                    }
                }
            }
        }

        if (exhibitingOrganisation.filterCategories[1]) {
            if (exhibitingOrganisation.filterCategories[1].responses && exhibitingOrganisation.filterCategories[1].responses.length > 0) {
                let result = exhibitingOrganisation.filterCategories[1].responses
                for (let i = 0; i < result.length; i++) {
                    let item = result[i].multilingual
                    target += item[1].name;
                    if (i < result.length - 1) {
                        target += '||';
                    }
                }
            }
        }

        let newData = [item.companyName, item.countryName, address, item.website, item.email, item.phone, item.description, product, target];

        addDataToExcelFile(newData);
        console.log("create success", stt)
    } catch (error) {
        console.error("Error fetching data:", error);
    }
};

// Hàm để thêm dữ liệu mới vào tệp Excel đã tồn tại
const addDataToExcelFile = (newData) => {
    // Đọc tệp Excel
    const workbook = XLSX.readFile(filename);

    // Lấy sheet cần thêm dữ liệu (ở đây tôi sẽ lấy sheet đầu tiên)
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Chuyển dữ liệu từ sheet thành mảng
    const dataArray = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Tạo trường STT cho dữ liệu mới và thêm vào newData
    const stt = dataArray.length + 1; // Lấy STT cho hàng mới
    const newDataWithSTT = [stt, ...newData]; // Thêm STT vào đầu dữ liệu mới

    // Thêm dữ liệu mới vào cuối mảng
    dataArray.push(newDataWithSTT);

    // Tạo lại sheet từ mảng dữ liệu đã được cập nhật
    const updatedSheet = XLSX.utils.aoa_to_sheet(dataArray);

    // Gán sheet mới vào workbook
    workbook.Sheets[sheetName] = updatedSheet;

    // Lưu lại tệp Excel
    XLSX.writeFile(workbook, filename);
};

app.get('/create', (req, res) => {

    if (checkFileExists(filename)) {
        res.send('File already exists');
    } else {

        const header = ["STT", "TÊN CÔNG TY", "QUỐC GIA", "ĐỊA CHỈ", "WEBSITE", "EMAIL", "SĐT", "MÔ TẢ", "DANH MỤC SẢN PHẨM", "KHÁCH HÀNG MỤC TIÊU"];

        // Tạo file Excel
        createExcelFile(header);
        res.send('Excel file created');
    }
});

app.get('/add', async (req, res) => {
    await fetchData(0)
    res.send('Success');
});

app.get('/add1', async (req, res) => {
    await fetchData(1)
    res.send('Success');
});

app.get('/add2', async (req, res) => {
    await fetchData(2)
    res.send('Success');
});

app.get('/add3', async (req, res) => {
    await fetchData(3)
    res.send('Success');
});

app.get('/add4', async (req, res) => {
    await fetchData(4)
    res.send('Success');
});

app.get('/add5', async (req, res) => {
    await fetchData(5)
    res.send('Success');
});

app.get('/add6', async (req, res) => {
    await fetchData(6)
    res.send('Success');
});


const PORT = 2000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});