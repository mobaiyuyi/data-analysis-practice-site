/* 研数通途 · 业务宇宙「云栖茶饮」
 * 所有练习共享同一家公司：12 家门店的连锁奶茶品牌。
 * 这里只放「名字」，不放数据——各题的工作簿数据直接写在题目对象里（见 spec §9.3）。
 */
(function (root) {
  'use strict';
  var Mobai = (root.Mobai = root.Mobai || {});

  var brand = '云栖茶饮';

  /** 12 家门店：全名 / 编号代码 / 所在区 / 区域经理 */
  var stores = [
    { name: '观山湖店', code: 'GSH', city: '观山湖区', manager: '陈立诚' },
    { name: '喷水池店', code: 'PSX', city: '云岩区', manager: '许安然' },
    { name: '花果园店', code: 'HGY', city: '南明区', manager: '罗景舟' },
    { name: '世纪城店', code: 'SJC', city: '观山湖区', manager: '陈立诚' },
    { name: '金融城店', code: 'JRC', city: '观山湖区', manager: '陈立诚' },
    { name: '会展城店', code: 'HZC', city: '观山湖区', manager: '陈立诚' },
    { name: '大学城店', code: 'DXC', city: '花溪区', manager: '顾南舟' },
    { name: '青岩古镇店', code: 'QYGZ', city: '花溪区', manager: '顾南舟' },
    { name: '白云公园店', code: 'BYGY', city: '白云区', manager: '傅熙和' },
    { name: '未来方舟店', code: 'WLFZ', city: '南明区', manager: '罗景舟' },
    { name: '龙洞堡店', code: 'LDB', city: '南明区', manager: '罗景舟' },
    { name: '小河转盘店', code: 'XHZP', city: '经开区', manager: '顾南舟' }
  ];

  /** 产品与标准单价（元） */
  var products = [
    { name: '茉莉云顶', price: 18 },
    { name: '生椰拿铁', price: 22 },
    { name: '杨枝甘露', price: 24 },
    { name: '黑糖珍珠', price: 20 },
    { name: '四季春茶', price: 16 },
    { name: '芝士葡萄', price: 26 },
    { name: '桂花乌龙', price: 21 },
    { name: '柠檬茶', price: 17 }
  ];

  /** 公司里的人 */
  var people = [
    { name: '林小满', role: '数据分析师', code: 'YQ0001' },
    { name: '周雨薇', role: '运营主管', code: 'YQ0002' },
    { name: '陈立诚', role: '区域经理（观山湖）', code: 'YQ0003' },
    { name: '许安然', role: '区域经理（云岩）', code: 'YQ0004' },
    { name: '罗景舟', role: '区域经理（南明）', code: 'YQ0005' },
    { name: '顾南舟', role: '区域经理（花溪·经开）', code: 'YQ0006' },
    { name: '傅熙和', role: '区域经理（白云）', code: 'YQ0007' }
  ];

  function storeNames() {
    return stores.map(function (s) { return brand + '·' + s.name; });
  }

  function fullName(shortName) {
    return brand + '·' + shortName;
  }

  function storeByCode(code) {
    for (var i = 0; i < stores.length; i++) {
      if (stores[i].code === code) return stores[i];
    }
    return null;
  }

  Mobai.Universe = {
    brand: brand,
    stores: stores,
    products: products,
    people: people,
    storeNames: storeNames,
    fullName: fullName,
    storeByCode: storeByCode,
    today: '2026-03-16',          // 全站固定「今天」，让日期题的期望值可复现
    currency: '元',
    dateHint: '日期一律写成 2026-03-09 这样的格式'
  };
})(typeof window !== 'undefined' ? window : globalThis);