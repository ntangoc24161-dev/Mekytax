const crypto = require('crypto');

function evpBytesToKey(password, salt, keyLen, ivLen) {
    let d = Buffer.alloc(0);
    let totalLen = keyLen + ivLen;
    let mdBuf = Buffer.alloc(0);
    while (d.length < totalLen) {
        const hash = crypto.createHash('md5');
        if (mdBuf.length > 0) hash.update(mdBuf);
        hash.update(password);
        hash.update(salt);
        mdBuf = hash.digest();
        d = Buffer.concat([d, mdBuf]);
    }
    return {
        key: d.subarray(0, keyLen),
        iv: d.subarray(keyLen, keyLen + ivLen)
    };
}

function dcV2(ciphertextB64, keyStr = "jcjioe92ujdkod021") {
    const raw = Buffer.from(ciphertextB64, 'base64');
    if (raw.subarray(0, 8).toString() !== 'Salted__') {
        throw new Error('Not salted OpenSSL format');
    }
    const salt = raw.subarray(8, 16);
    const encData = raw.subarray(16);
    const { key, iv } = evpBytesToKey(Buffer.from(keyStr, 'utf8'), salt, 32, 16);
    const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
    let decrypted = Buffer.concat([decipher.update(encData), decipher.final()]);
    return JSON.parse(decrypted.toString('utf8'));
}

function ecV2(dataObj, keyStr = "jcjioe92ujdkod021") {
    const salt = crypto.randomBytes(8);
    const { key, iv } = evpBytesToKey(Buffer.from(keyStr, 'utf8'), salt, 32, 16);
    const cipher = crypto.createCipheriv('aes-256-cbc', key, iv);
    const plaintext = JSON.stringify(dataObj);
    const encData = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
    const salted = Buffer.concat([Buffer.from('Salted__', 'utf8'), salt, encData]);
    return salted.toString('base64');
}

const action = process.argv[2];
const payload = process.argv[3];

if (action === 'decrypt' && payload) {
    try {
        const res = dcV2(payload);
        process.stdout.write(JSON.stringify(res));
    } catch (e) {
        process.stderr.write("Decrypt error: " + e.message);
        process.exit(1);
    }
} else if (action === 'encrypt' && payload) {
    try {
        const obj = JSON.parse(payload);
        const res = ecV2(obj);
        process.stdout.write(res);
    } catch (e) {
        process.stderr.write("Encrypt error: " + e.message);
        process.exit(1);
    }
}

module.exports = { dcV2, ecV2, evpBytesToKey };
