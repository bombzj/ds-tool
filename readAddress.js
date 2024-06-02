const fs = require('fs');
const mysql = require('mysql');
const path = require('path');
const xlsx = require('xlsx');


// read address list from aaa.xslx, header: countryId	countrySelf	countryName	countryArName	stateId	stateSelf	stateName	stateArName	cityId	citySelf	cityName	cityArName	districtId	districtSelf	districtName	districtArName
// connect to mysql: localhost, root/1qaz2wsx, db:qshop
// convert address and save to table addr_city, column: id, code, name, parent_id, level, country, seq, lang, local_name, name_en

// Read the address list from aaa.xlsx
const workbook = xlsx.readFile('aaa.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const addressList = xlsx.utils.sheet_to_json(worksheet);

// Connect to MySQL
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '1QAZ2wsx',
    database: 'qshop'
});

function querySync(query, values) {
    return new Promise((resolve, reject) => {
        connection.query(query, values, (error, results) => {
            if (error) reject(error);
            resolve(results);
        });
    });
}

async function doit()
{

    connection.connect();

    let addrCountry = await querySync('select * from addr_country', []);
    let countryMap = new Map();
    for(let i = 0; i < addrCountry.length; i++) {
        const { id, code, name } = addrCountry[i];
        countryMap.set(name, {id, code});
    }

    let visited = new Map();

    // Convert and save address to table addr_city, sync
    for(let i = 0; i < addressList.length; i++) {
        let { countryId, countrySelf, countryName, countryArName, stateId, stateSelf, stateName, stateArName, cityId, citySelf, cityName, cityArName, districtId, districtSelf, districtName, districtArName } = addressList[i];
        // if(i > 100) break;
        if(!countryMap.has(countryName)) {
            continue;
        }
        const countryData = countryMap.get(countryName);
        if(visited.has(stateSelf)) {
            stateId = visited.get(stateSelf);
        } else {
            const query = `INSERT INTO addr_city (code, name, parent_id, level, country, country_id, seq, lang, local_name, name_en) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
            if(!stateArName) stateArName = stateName;
            const values = [stateSelf, stateName, 0, 1, countryData.code, countryData.id, stateSelf, '', stateArName, stateName];
            const res = await querySync(query, values);
            // console.log('Address saved:', res);
            stateId = res.insertId;
            visited.set(stateSelf, stateId);
        }

        if(visited.has(citySelf)) {
            cityId = visited.get(citySelf);
        } else {
            const query = `INSERT INTO addr_city (code, name, parent_id, level, country, country_id, seq, lang, local_name, name_en) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
            if(!cityArName) cityArName = cityName;
            const values = [citySelf, cityName, stateId, 2, countryData.code, countryData.id, citySelf, '', cityArName, cityName];
            const res = await querySync(query, values);
            // console.log('Address saved:', res);
            cityId = res.insertId;
            visited.set(citySelf, cityId);
        }
        if(districtName && districtName !== '') {
            const query = `INSERT INTO addr_city (code, name, parent_id, level, country, country_id, seq, lang, local_name, name_en) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
            if(!districtArName) districtArName = districtName;
            const values = [districtSelf, districtName, cityId, 3, countryData.code, countryData.id, districtSelf, '', districtArName, districtName];
            const res = await querySync(query, values);
            // console.log('Address saved:', res);
        }

    }

    connection.end();
}

doit();