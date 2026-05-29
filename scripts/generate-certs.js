/**
 * 生成本地 HTTPS 自签名证书
 * 使用 node-forge 库，纯 JavaScript 实现，无外部依赖
 */
const forge = require('node-forge');
const fs = require('fs');
const path = require('path');

const CERTS_DIR = path.join(__dirname, '..', 'certs');

function main() {
  // 确保 certs 目录存在
  if (!fs.existsSync(CERTS_DIR)) {
    fs.mkdirSync(CERTS_DIR, { recursive: true });
  }

  const certPath = path.join(CERTS_DIR, 'localhost.crt');
  const keyPath = path.join(CERTS_DIR, 'localhost.key');

  // 如果证书已存在，询问是否覆盖
  if (fs.existsSync(certPath)) {
    // 在自动化场景下，如果证书存在就直接返回（不报错）
    if (process.argv.includes('--force')) {
      console.log('检测到已有证书，使用 --force 强制重新生成');
    } else {
      console.log('[OK] 证书文件已存在，跳过生成');
      console.log('    如需重新生成，请使用: node scripts/generate-certs.js --force');
      return;
    }
  }

  console.log('正在生成 RSA 密钥对 (2048 bit)...');

  // 1. 生成 RSA 密钥对
  const keys = forge.pki.rsa.generateKeyPair(2048);

  // 2. 创建证书
  const cert = forge.pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = '01' + Date.now().toString(16);

  // 证书有效期：从当前时间起 10 年
  const now = new Date();
  cert.validity.notBefore = now;
  cert.validity.notAfter = new Date(now.getFullYear() + 10, now.getMonth(), now.getDate());

  // 证书主体和颁发者（自签名，二者相同）
  const attrs = [
    { name: 'commonName', value: 'localhost' },
    { name: 'countryName', value: 'CN' },
    { name: 'organizationName', value: 'M365 AI Add-in Local Dev' }
  ];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);

  // 添加 SAN（Subject Alternative Name）扩展
  cert.setExtensions([
    {
      name: 'basicConstraints',
      cA: true,
      critical: true
    },
    {
      name: 'keyUsage',
      keyCertSign: true,
      digitalSignature: true,
      nonRepudiation: true,
      keyEncipherment: true,
      dataEncipherment: true,
      critical: true
    },
    {
      name: 'extKeyUsage',
      serverAuth: true,
      clientAuth: true
    },
    {
      name: 'subjectAltName',
      altNames: [
        { type: 2, value: 'localhost' },   // DNS
        { type: 7, ip: '127.0.0.1' },      // IP
        { type: 7, ip: '::1' }             // IPv6
      ]
    }
  ]);

  // 4. 用私钥签名证书
  cert.sign(keys.privateKey, forge.md.sha256.create());

  // 5. 导出 PEM 格式
  const certPem = forge.pki.certificateToPem(cert);
  const keyPem = forge.pki.privateKeyToPem(keys.privateKey);

  // 6. 写入文件
  fs.writeFileSync(certPath, certPem, 'utf8');
  fs.writeFileSync(keyPath, keyPem, 'utf8');

  console.log('[OK] 证书生成完成！');
  console.log(`    证书: ${certPath}`);
  console.log(`    私钥: ${keyPath}`);

  // 7. 打印根证书安装提示
  console.log('');
  console.log('========================================');
  console.log('  重要：安装根证书到系统信任链');
  console.log('========================================');
  console.log('');
  console.log('请以管理员身份运行以下命令：');
  console.log(`  certutil -addstore -user Root "${certPath}"`);
  console.log('');
  console.log('或手动操作：');
  console.log('  1. 双击 certs/localhost.crt');
  console.log('  2. 点击"安装证书"');
  console.log('  3. 选择"本地计算机" → "将所有证书放入下列存储"');
  console.log('  4. 浏览 → "受信任的根证书颁发机构" → 确定');
  console.log('========================================');
}

main();
