// Flavor intensity chart for the Sichuan cooking guide
(function () {
  var root = document.documentElement;
  var style = getComputedStyle(root);
  var v = function (name) { return style.getPropertyValue(name).trim(); };

  var accent = v('--accent');
  var accent2 = v('--accent2');
  var accent3 = v('--accent3');
  var accent4 = v('--accent4');
  var accent5 = v('--accent5');
  var ink = v('--ink');
  var muted = v('--muted');
  var rule = v('--rule');
  var bg = v('--bg');
  var bg2 = v('--bg2');

  var el = document.getElementById('chart-flavor');
  if (!el || !window.echarts) return;

  var dims = ['麻', '辣', '酸甜', '咸鲜'];
  var series = [
    { name: '麻婆豆腐', data: [5, 4, 0.5, 4.5], color: accent },
    { name: '回锅肉', data: [1, 3, 1, 5], color: accent3 },
    { name: '宫保鸡丁', data: [2, 3, 4.5, 3.5], color: accent2 },
    { name: '鱼香肉丝', data: [0.5, 3, 5, 4], color: accent5 },
    { name: '水煮肉片', data: [4.5, 5, 0, 4], color: accent4 },
    { name: '干煸四季豆', data: [2.5, 3.5, 0.5, 4], color: muted }
  ];

  var chart = echarts.init(el, null, { renderer: 'svg' });

  chart.setOption({
    animation: false,
    backgroundColor: 'transparent',
    textStyle: {
      fontFamily: "'WorkSans','PingFang SC','Microsoft YaHei',sans-serif",
      color: ink
    },
    tooltip: {
      trigger: 'axis',
      appendToBody: true,
      backgroundColor: bg,
      borderColor: rule,
      textStyle: { color: ink, fontSize: 12 },
      axisPointer: { type: 'shadow', shadowStyle: { color: 'rgba(43,30,23,0.04)' } }
    },
    legend: {
      bottom: 0,
      itemWidth: 10,
      itemHeight: 10,
      itemGap: 14,
      textStyle: { color: muted, fontSize: 11 }
    },
    grid: { left: 40, right: 12, top: 18, bottom: 62 },
    xAxis: {
      type: 'category',
      data: dims,
      axisLabel: { color: ink, fontSize: 12 },
      axisTick: { show: false },
      axisLine: { lineStyle: { color: rule } }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 5,
      interval: 1,
      name: '强度',
      nameTextStyle: { color: muted, fontSize: 11 },
      axisLabel: { color: muted, fontSize: 11 },
      splitLine: { lineStyle: { color: rule, type: 'dashed' } }
    },
    series: series.map(function (s) {
      return {
        name: s.name,
        type: 'bar',
        data: s.data,
        barMaxWidth: 14,
        itemStyle: { color: s.color, borderRadius: [2, 2, 0, 0] },
        emphasis: { itemStyle: { color: s.color } }
      };
    })
  });

  window.addEventListener('resize', function () { chart.resize(); });
})();
