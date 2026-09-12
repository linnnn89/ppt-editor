export function bicycle() {
  const items = [];
  const line = (name, x1, y1, x2, y2, color = 'E05244', weight = 8) => {
    // Canonical endpoints run left to right, making flipV independently testable.
    if (x1 > x2 || (x1 === x2 && y1 > y2)) [x1, y1, x2, y2] = [x2, y2, x1, y1];
    items.push({ name, type: 'line', x1, y1, x2, y2, left: x1, top: Math.min(y1, y2), width: x2 - x1, height: Math.abs(y2 - y1), color, weight });
  };
  const circle = (name, x, y, radius, color, weight, fill) => items.push({ name, type: 'ellipse', left: x - radius, top: y - radius, width: radius * 2, height: radius * 2, color, weight, ...(fill ? { fill } : {}) });
  line('Ground', 145, 490, 830, 490, 'DCE2E8', 2);
  for (const [label, x] of [['Rear', 290], ['Front', 690]]) {
    circle(`${label}_Tire`, x, 370, 112, '263442', 9);
    circle(`${label}_Rim`, x, 370, 102, '9BA9B7', 2);
    for (let n = 0; n < 12; n++) {
      const a = n * Math.PI / 6;
      line(`${label}_Spoke_${n + 1}`, x, 370, x + Math.cos(a) * 100, 370 + Math.sin(a) * 100, 'B5C0CB', 1);
    }
  }
  line('Rear_Stay', 290, 370, 410, 215);
  line('Chain_Stay', 290, 370, 470, 370);
  line('Seat_Tube', 410, 215, 470, 370);
  line('Top_Tube', 410, 215, 622, 215);
  line('Down_Tube', 470, 370, 622, 215);
  line('Fork', 622, 215, 690, 370, '34485C', 8);
  line('Seat_Post', 397, 181, 410, 215, '34485C', 7);
  line('Saddle', 360, 177, 440, 177, '263442', 13);
  line('Handle_Stem', 603, 165, 622, 215, '34485C', 7);
  line('Handlebar', 603, 165, 658, 165, '34485C', 7);
  line('Handle_Drop', 658, 165, 674, 183, '34485C', 7);
  line('Handle_Grip', 656, 190, 674, 183, '263442', 9);
  line('Chain_Upper', 290, 365, 470, 350, '687889', 1.5);
  line('Chain_Lower', 290, 375, 470, 390, '687889', 1.5);
  circle('Chainring', 470, 370, 21, '34485C', 4, 'FFFFFF');
  line('Crank', 450, 346, 491, 396, '34485C', 5);
  line('Pedal_Back', 439, 345, 460, 345, '263442', 5);
  line('Pedal_Front', 481, 397, 504, 397, '263442', 5);
  circle('Rear_Hub', 290, 370, 6, '34485C', 1, '34485C');
  circle('Front_Hub', 690, 370, 6, '34485C', 1, '34485C');
  circle('Crank_Axle', 470, 370, 4, 'E05244', 1, 'E05244');
  return items;
}

export function bicycleDeck(items = bicycle()) {
  return { title: 'Editable bicycle', width: 960, height: 540, slides: [{ background: 'FFFFFF', items: items.map(i => ({ type: 'shape', shape: i.type, name: i.name, left: i.left, top: i.top, width: i.width, height: i.height, line: i.color, lineWidth: i.weight, ...(i.type === 'line' && i.y2 < i.y1 ? { flipV: true } : {}), ...(i.fill ? { fill: i.fill } : {}) })) }] };
}
