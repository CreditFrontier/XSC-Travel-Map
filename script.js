const chart = echarts.init(document.getElementById('map'));
const MAP_BASE = 'https://cdn.jsdelivr.net/npm/echarts@4.9.0/map/json/';
const STORAGE_KEY = 'travel-footprint-v1';
const EXCEL_IMPORT_KEY = 'travel-footprint-excel-v3-city-imported';
const SEP = '::';

// 来自《旅游景点汇总_V3完整版.xlsx》，只保留能明确归属的地级市/自治州。
// 直辖市没有地级市，按 Excel 明确记录的区县点亮；“各地”等模糊记录不导入。
const EXCEL_VISITED = [
  ['安徽', '黄山市'],
  ['甘肃', '白银市'], ['甘肃', '甘南藏族自治州'], ['甘肃', '嘉峪关市'],
  ['甘肃', '酒泉市'], ['甘肃', '兰州市'], ['甘肃', '临夏回族自治州'],
  ['甘肃', '陇南市'], ['甘肃', '天水市'], ['甘肃', '张掖市'],
  ['广东', '广州市'], ['广东', '珠海市'],
  ['广西', '北海市'], ['广西', '崇左市'], ['广西', '桂林市'],
  ['广西', '柳州市'], ['广西', '南宁市'],
  ['贵州', '安顺市'], ['贵州', '贵阳市'], ['贵州', '六盘水市'],
  ['贵州', '黔东南苗族侗族自治州'], ['贵州', '黔南布依族苗族自治州'],
  ['贵州', '铜仁市'], ['贵州', '遵义市'],
  ['河北', '邯郸市'], ['河北', '邢台市'],
  ['河南', '焦作市'], ['河南', '洛阳市'], ['河南', '三门峡市'], ['河南', '新乡市'],
  ['湖北', '恩施土家族苗族自治州'], ['湖北', '神农架林区'], ['湖北', '十堰市'],
  ['湖北', '武汉市'], ['湖北', '宜昌市'],
  ['湖南', '衡阳市'], ['湖南', '湘潭市'], ['湖南', '岳阳市'], ['湖南', '长沙市'],
  ['江苏', '苏州市'], ['江苏', '盐城市'], ['江苏', '镇江市'],
  ['江西', '上饶市'],
  ['宁夏', '银川市'], ['宁夏', '吴忠市'],
  ['青海', '海北藏族自治州'], ['青海', '海南藏族自治州'], ['青海', '西宁市'],
  ['山东', '枣庄市'],
  ['山西', '大同市'], ['山西', '朔州市'], ['山西', '晋中市'], ['山西', '临汾市'],
  ['山西', '太原市'], ['山西', '忻州市'], ['山西', '运城市'],
  ['陕西', '延安市'], ['陕西', '榆林市'],
  ['上海', '宝山区'], ['上海', '虹口区'], ['上海', '黄浦区'], ['上海', '嘉定区'],
  ['上海', '静安区'], ['上海', '闵行区'], ['上海', '浦东新区'], ['上海', '普陀区'],
  ['上海', '青浦区'], ['上海', '松江区'], ['上海', '徐汇区'], ['上海', '长宁区'],
  ['四川', '阿坝藏族羌族自治州'],
  ['浙江', '杭州市'], ['浙江', '湖州市'], ['浙江', '金华市'],
  ['浙江', '宁波市'], ['浙江', '衢州市'],
  ['重庆', '大足区'], ['重庆', '武隆区']
].map(([province, city]) => province + SEP + city);

const OLD_EXCEL_PROVINCES = new Set([
  '河北', '山西', '上海', '江苏', '浙江', '安徽', '江西', '山东', '河南',
  '湖北', '湖南', '广东', '广西', '重庆', '四川', '贵州', '西藏', '陕西',
  '甘肃', '青海', '宁夏', '新疆', '香港'
].map(province => province + SEP + '__ALL__'));

const PINYIN = {
  '北京': 'beijing', '天津': 'tianjin', '河北': 'hebei', '山西': 'shanxi',
  '内蒙古': 'neimenggu', '辽宁': 'liaoning', '吉林': 'jilin', '黑龙江': 'heilongjiang',
  '上海': 'shanghai', '江苏': 'jiangsu', '浙江': 'zhejiang', '安徽': 'anhui',
  '福建': 'fujian', '江西': 'jiangxi', '山东': 'shandong', '河南': 'henan',
  '湖北': 'hubei', '湖南': 'hunan', '广东': 'guangdong', '广西': 'guangxi',
  '海南': 'hainan', '重庆': 'chongqing', '四川': 'sichuan', '贵州': 'guizhou',
  '云南': 'yunnan', '西藏': 'xizang', '陕西': 'shanxi1', '甘肃': 'gansu',
  '青海': 'qinghai', '宁夏': 'ningxia', '新疆': 'xinjiang', '台湾': 'taiwan',
  '香港': 'xianggang', '澳门': 'aomen'
};
const ALT = { '陕西': ['shaanxi'], '香港': ['hongkong'], '澳门': ['macau'] };

let level = 'china';
let currentSubMapName = null;
const visited = new Set(loadSaved());
const loadingEl = document.getElementById('loading');
const msgEl = document.getElementById('msg');

function loadSaved() {
  try {
    const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    const saved = Array.isArray(value)
      ? value.filter(item => typeof item === 'string' && !OLD_EXCEL_PROVINCES.has(item))
      : [];
    if (localStorage.getItem(EXCEL_IMPORT_KEY) !== '1') {
      const merged = [...new Set([...saved, ...EXCEL_VISITED])];
      localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
      localStorage.setItem(EXCEL_IMPORT_KEY, '1');
      return merged;
    }
    return saved;
  } catch (_) { return []; }
}
function saveVisited() { localStorage.setItem(STORAGE_KEY, JSON.stringify([...visited])); }
function showLoading(show) { loadingEl.style.display = show ? 'flex' : 'none'; }
function showMsg(text) { msgEl.textContent = text; msgEl.style.display = text ? 'block' : 'none'; }
function provCount(province) {
  let count = 0;
  visited.forEach(key => { if (key.startsWith(province + SEP)) count += 1; });
  return count;
}

async function fetchGeo(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`地图数据加载失败 (${response.status})`);
  return response.json();
}
async function loadChina() {
  if (!echarts.getMap('china')) echarts.registerMap('china', await fetchGeo(MAP_BASE + 'china.json'));
}
async function loadProvince(province) {
  const candidates = [PINYIN[province], ...(ALT[province] || [])].filter(Boolean);
  for (const name of candidates) {
    if (echarts.getMap(name)) return name;
    try {
      echarts.registerMap(name, await fetchGeo(`${MAP_BASE}province/${name}.json`));
      return name;
    } catch (_) { /* 尝试备用文件名 */ }
  }
  return null;
}

function baseOption(mapName, data, maxValue) {
  const isHainan = mapName === 'hainan';
  return {
    tooltip: {
      trigger: 'item',
      formatter: params => `${params.name}<br>${Number(params.value) > 0 ? '✓ 已标记' : '未标记'}`
    },
    visualMap: {
      min: 0, max: Math.max(1, maxValue), left: 12, bottom: 12, orient: 'horizontal',
      text: ['多', '未去'], calculable: false,
      inRange: { color: ['#eef2f7', '#7ee0c9', '#14b8a6', '#0f766e'] },
      itemWidth: 12, itemHeight: 90, textStyle: { fontSize: 11 }
    },
    series: [{
      type: 'map', map: mapName, roam: true,
      // 海南数据包含远离本岛的南海岛屿；单独聚焦本岛，避免地图被压缩得像空白。
      center: isHainan ? [109.8, 19.2] : undefined,
      zoom: mapName === 'china' ? 1.15 : (isHainan ? 4.5 : 1),
      label: { show: false },
      emphasis: { label: { show: true, fontSize: 12 }, itemStyle: { areaColor: '#fcd34d' } },
      itemStyle: { borderColor: '#fff', borderWidth: 0.6 }, data
    }]
  };
}

async function renderChina() {
  showLoading(true); showMsg('');
  try {
    await loadChina();
    level = 'china'; currentSubMapName = null;
    const geo = echarts.getMap('china').geoJSON;
    let maxValue = 1;
    const data = geo.features.map(feature => {
      const name = feature.properties.name;
      const value = provCount(name);
      maxValue = Math.max(maxValue, value);
      return { name, value };
    });
    chart.setOption(baseOption('china', data, maxValue), true);
    document.getElementById('btn-back').style.display = 'none';
    document.getElementById('btn-markall').style.display = 'none';
    document.getElementById('crumb').innerHTML = '<b>全国</b>';
    updateStats();
  } catch (error) {
    showMsg('全国地图加载失败。请确认设备已联网，然后刷新页面。');
  } finally { showLoading(false); }
}

async function renderProvince(province) {
  showLoading(true); showMsg('');
  const mapName = await loadProvince(province);
  level = province;
  document.getElementById('btn-back').style.display = '';
  document.getElementById('btn-markall').style.display = '';
  document.getElementById('crumb').innerHTML = `<span>全国 / </span><b>${province}</b>`;
  if (!mapName) {
    currentSubMapName = null;
    showLoading(false);
    showMsg(`“${province}”的下级地图暂时无法加载，可用“标记整省已到访”记录该省。`);
    return;
  }
  currentSubMapName = mapName;
  const geo = echarts.getMap(mapName).geoJSON;
  const allMarked = visited.has(province + SEP + '__ALL__');
  const data = geo.features.map(feature => ({
    name: feature.properties.name,
    value: allMarked || visited.has(province + SEP + feature.properties.name) ? 1 : 0
  }));
  chart.setOption(baseOption(mapName, data, 1), true);
  showLoading(false);
  updateStats();
}

chart.on('click', params => {
  if (params.componentType !== 'series') return;
  if (level === 'china') {
    if (PINYIN[params.name]) renderProvince(params.name);
    return;
  }
  const key = level + SEP + params.name;
  visited.has(key) ? visited.delete(key) : visited.add(key);
  saveVisited();
  renderProvince(level);
});

document.getElementById('btn-back').onclick = renderChina;
document.getElementById('btn-markall').onclick = () => {
  const key = level + SEP + '__ALL__';
  visited.has(key) ? visited.delete(key) : visited.add(key);
  saveVisited();
  currentSubMapName ? renderProvince(level) : updateStats();
};

function updateStats() {
  const provinces = new Set([...visited].map(key => key.split(SEP)[0]));
  const subCount = [...visited].filter(key => !key.endsWith('__ALL__')).length;
  document.getElementById('s-prov').textContent = provinces.size;
  document.getElementById('s-sub').textContent = subCount;
  document.getElementById('s-rate').textContent = Math.round(provinces.size / 34 * 100) + '%';
  const listEl = document.getElementById('visited-list');
  if (!visited.size) {
    listEl.innerHTML = '<span class="empty">还没有记录，去点亮第一个地方吧！</span>';
    return;
  }
  const groups = {};
  visited.forEach(key => {
    const [province, area] = key.split(SEP);
    (groups[province] ||= []).push(area);
  });
  listEl.replaceChildren();
  Object.keys(groups).sort().forEach(province => {
    const group = document.createElement('div');
    group.className = 'visited-group';
    const title = document.createElement('b');
    title.textContent = province;
    group.append(title, document.createElement('br'));
    groups[province].forEach(area => {
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.append(document.createTextNode(area === '__ALL__' ? '整省' : area));
      const remove = document.createElement('i');
      remove.textContent = '×'; remove.title = '删除'; remove.tabIndex = 0;
      remove.onclick = () => {
        visited.delete(province + SEP + area); saveVisited();
        level === 'china' ? renderChina() : (currentSubMapName ? renderProvince(level) : updateStats());
      };
      remove.onkeydown = event => { if (event.key === 'Enter' || event.key === ' ') remove.click(); };
      tag.append(remove); group.append(tag);
    });
    listEl.append(group);
  });
}

const modalBg = document.getElementById('modal-bg');
const modalText = document.getElementById('modal-text');
const modalTitle = document.getElementById('modal-title');
const modalOk = document.getElementById('modal-ok');
document.getElementById('btn-export').onclick = () => {
  modalTitle.textContent = '导出记录'; modalText.value = JSON.stringify([...visited]);
  modalOk.style.display = 'none'; modalBg.style.display = 'flex'; modalText.select();
};
document.getElementById('btn-import').onclick = () => {
  modalTitle.textContent = '导入记录'; modalText.value = '';
  modalOk.style.display = ''; modalBg.style.display = 'flex'; modalText.focus();
};
document.getElementById('modal-cancel').onclick = () => { modalBg.style.display = 'none'; };
modalOk.onclick = () => {
  try {
    const records = JSON.parse(modalText.value);
    if (!Array.isArray(records) || !records.every(item => typeof item === 'string')) throw new Error();
    visited.clear(); records.forEach(item => visited.add(item)); saveVisited();
    modalBg.style.display = 'none'; renderChina();
  } catch (_) { showMsg('导入失败：请粘贴此前导出的完整记录文本。'); }
};
modalBg.onclick = event => { if (event.target === modalBg) modalBg.style.display = 'none'; };
window.addEventListener('keydown', event => { if (event.key === 'Escape') modalBg.style.display = 'none'; });
window.addEventListener('resize', () => chart.resize());

renderChina();
