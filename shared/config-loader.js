/**
 * 配置加载模块
 * 读取本地 config.json 文件
 */

let cachedConfig = null;

/**
 * 获取配置文件的 URL
 * 相对于当前页面的位置自动推导，不硬编码端口
 */
function getConfigUrl() {
  // 优先从当前页面 origin 推导
  const base = window.location.origin;
  // 如果页面是通过 office-addin-debugging 加载的，它可能在另一个端口
  // 我们先尝试同源加载，失败后回退到 127.0.0.1:3000
  return base + '/config/config.json';
}

/**
 * 读取本地 config.json
 * 若 apiKey 为空或为默认占位文本，抛出友好的中文提示
 * @returns {Promise<object>} 配置对象
 */
async function loadConfig() {
  const isPlaceholder = (key) => !key || key === 'placeholder' || key === '在此填入你的 DeepSeek API Key';

  if (cachedConfig && cachedConfig.apiKey && !isPlaceholder(cachedConfig.apiKey)) {
    return cachedConfig;
  }

  const urls = [
    getConfigUrl(),
    'https://127.0.0.1:3000/config/config.json',
    'https://localhost:3000/config/config.json'
  ];

  let lastError = null;

  for (const url of urls) {
    try {
      const response = await fetch(url);

      if (!response.ok) {
        throw new Error('HTTP ' + response.status);
      }

      const config = await response.json();

      if (isPlaceholder(config.apiKey)) {
        throw new Error(
          '请先在 config/config.json 中填写你的 DeepSeek API Key，然后重启插件服务。\n' +
          '获取 Key: https://platform.deepseek.com/api_keys'
        );
      }

      cachedConfig = config;
      return config;

    } catch (error) {
      lastError = error;
      // 继续尝试下一个 URL
    }
  }

  // 所有 URL 都失败
  throw new Error(
    '无法加载配置文件。请确保已运行 start.bat 启动服务。\n' +
    (lastError ? '（' + lastError.message + '）' : '')
  );
}

/**
 * 清除配置缓存
 */
function clearConfigCache() {
  cachedConfig = null;
}

/**
 * 获取指定应用的模型名称
 * @param {string} appType - 应用类型（word, excel, ppt）
 * @returns {Promise<string>} 模型名称
 */
async function getModelForApp(appType) {
  const config = await loadConfig();
  return config.models[appType] || 'deepseek-v4-flash';
}
