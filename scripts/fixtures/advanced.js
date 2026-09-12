const title = text => ({ type: 'text', name: 'Title', text, left: 40, top: 28, width: 880, height: 60, fontSize: 30, bold: true });
const chart = (chartType, name, series, left = 40, width = 420) => ({ type: 'chart', chartType, name, series, left, top: 130, width, height: 330 });
const series = [{ name: 'Group A', labels: ['Q1','Q2','Q3','Q4'], values: [18,24,31,39] }, { name: 'Group B', labels: ['Q1','Q2','Q3','Q4'], values: [12,18,22,29] }];
export function advancedDeck() {
  return { title: 'Advanced editable objects — synthetic demonstration', theme: 'executive', slides: [
    { items: [title('01  |  Volume and rate / 双轴组合图'), chart('combo','Combo', [{ ...series[0], type: 'bar' }, { name: 'Rate (%)', labels: series[0].labels, values: [42,55,68,74], type: 'line', secondaryAxis: true }], 40, 600),
      { type: 'text', name: 'Insight', text: 'READING THE CHART\n\nBars: volume\nLine: rate (%)\n\nSynthetic data only', left: 690, top: 150, width: 220, height: 290, fontSize: 22 } ] },
    { items: [title('02  |  Comparison and trends / 比较与趋势'), { ...chart('bar','Stacked',series), stacked: true }, chart('line','Trend',series,500) ] },
    { items: [title('03  |  Share and profile / 占比与轮廓'), chart('doughnut','Share',[series[0]]), chart('radar','Profile',series,500) ] },
    { items: [title('04  |  Accumulation / 面积图'), { ...chart('area','Area',series,40,880), stacked: true } ] },
    ...['process','cycle','hierarchy'].map((layout, index) => ({ items: [title(`0${index + 5}  |  Native SmartArt / Hierarchy example ${index + 1}`),
      { type:'text', name: `SmartArt_${layout}`, text: layout === 'hierarchy' ? 'Project lead\nAnalysis\nDelivery' : 'Discover\nDevelop\nReview', left:80,top:150,width:800,height:300,fontSize:26 } ] })),
    { items: [title('08  |  Visual reflow / 内容保留与视觉重排'),
      { type:'text',name:'Card1',text:'01  CLEAR MESSAGE\nOne idea per panel.',left:80,top:100,width:400,height:140,fontSize:22 },
      { type:'text',name:'Card2',text:'02  KEEP THE DATA\nNo values or text are removed.',left:170,top:150,width:450,height:120,fontSize:22 },
      { ...chart('bar','LayoutChart',[series[0]],300,500),top:180,height:250 },
      { type:'shape',shape:'roundRect',name:'Card3',text:'03  REVIEW OUTPUT\nInspect the final slide.',left:380,top:270,width:400,height:140,fill:'DCECEF',fontSize:22 } ] }
  ] };
}
