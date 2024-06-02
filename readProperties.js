const fs = require('fs');
const mysql = require('mysql');
const properties = require('java-properties');

// Create a MySQL connection
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'qshop',
    password: '1qaz2wsx',
    database: 'qshop'
});

async function readPropertiesFile(path) {
    // iterate through all files in path, file names are like: common_en.properties
    fs.readdir(path, (err, files) => {
        if (err) {
            console.error(err);
            return;
        }

        files.forEach(file => {
            [name, lang] = file.split('_');
            if (lang === undefined || !lang.endsWith('.properties')) {
                return;
            }
            // remove ext from lang
            lang = lang.split('.')[0];
            var values = properties.of(path + '/' + file);
            // save to i18n_message (code, message, locale) values ( key, value, lang )
            for (const key in values.objs) {
                connection.query('INSERT INTO i18n_message (code, message, locale) VALUES (?, ?, ?)', [key, values.objs[key], lang], (err, result) => {
                    if (err) {
                        console.error(err);
                        return;
                    }
                    console.log('Inserted: ' + key + ' - ' + values.objs[key] + ' - ' + lang);
                });
            }

 
        });
    });
}


readPropertiesFile('D:/git/qshop/qshop/src/main/resources/i18n');