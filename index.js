
function svgCircle(x, y, r) {
  return ` <circle cx="${x}" cy="${y}" r="${r}"/>`;
}

function svgRect(x, y, width, height) {
  return `<rect x="${x}" y="${y}" width="${width}" height="${height}" style="fill:none;stroke-width:3;stroke:rgb(0,0,0)" />`
}


/////////////////////////////////////
class CxyzAtt {
  x =[];
  y =[];
  z =[];
  constructor() {
  }

  push(aX, aY, aZ) {
    this.x.push(aX);
    this.y.push(aY);
    this.z.push(aZ);
  };

  translate(dx, dy, dz) {
    for (let i=0; i< this.x.length; i++) {
      this.x[i] = this.x[i] + dx;
      this.y[i] = this.y[i] + dy;
      this.z[i] = this.z[i] + dz;
    }
  }

  scale(sx, sy, sz) {
    for (let i=0; i< this.x.length; i++) {
      this.x[i] = this.x[i] * sx;
      this.y[i] = this.y[i] * sy;
      this.z[i] = this.z[i] * sz;
    }    
  }
}

class CxyArr {
  x =[];
  y =[];
  lineColor="black";
  lineWidth = 1;
  constructor() {
  }

  push(aX, aY) {
    this.x.push(aX);
    this.y.push(aY);
  };

  translate(dx, dy) {
    for (let i=0; i< this.x.length; i++) {
      this.x[i] = this.x[i] + dx;
      this.y[i] = this.y[i] + dy;
    }
  }

  scale(sx, sy) {
    for (let i=0; i< this.x.length; i++) {
      this.x[i] = this.x[i] * sx;
      this.y[i] = this.y[i] * sy;
    }    
  }

  rotate(angle) {
    // alpha - degree
    const alphaRad = Math.PI/180*angle;
    const sinA = Math.sin(alphaRad);
    const cosA = Math.cos(alphaRad);
    

    for (let i=0; i< this.x.length; i++) {
      let x = this.x[i] * cosA - this.y[i] * sinA;
      let y = this.x[i] * sinA + this.y[i] * cosA;
      this.x[i] = x;
      this.y[i] = y;
    }    
  }

  get svgPoly() {
    let s ="" ;
    //let color = 'red';
    for (let i=0; i< this.x.length; i++) {
      s = s+" "+`${this.x[i]},${this.y[i]}`;
    }
    let sp =  `<polyline points="${s}"  style="fill:none;stroke:${this.lineColor};stroke-width:${this.lineWidth}" />`;
    return sp;
  }

  set lineColor(aColor) {
    this.lineColor = aColor;    
  }

  set lineWidth(aWidth) {
    this.lineWidth = aWidth;    
  }  
}

aXY =  new CxyArr;
aXY.push(0, 0);
aXY.push(100, 100);

for (i=0; i<10; i++) {
  aXY.push(Math.random()*400, Math.random()*300);
}

let s = aXY.svgPoly;

//aXY.translate(300, 300);
//aXY.scale(0.5, 0.5);
aXY.rotate(30);

aXY.lineColor = 'red';
aXY.lineWidth = 5;

let s1 = aXY.svgPoly;

//console.log(s);

svgElm = document.querySelector(".js-svg1");
console.log(svgElm); // 
let w = svgElm.clientWidth;
let sr = svgRect(0,0,svgElm.clientWidth,svgElm.clientHeight);
//console.log(sr);  
svgElm.innerHTML= sr+ svgCircle(40, 40, 30) + svgCircle(60, 60, 20) + s +s1 ;

//let c = makeCoord(10, 20,30);
//let c = mcGetCoordDeg(10, 20, 5, 7, 30);
//console.log(c);


const message = document.getElementById('message');

//  GET the value of textarea
console.log(message.value); //  ""

// --------------------------------------

//  SET the value of textarea
message.value = '0 0 0 \n900 0 0 \n1200 10 5';


//  Append to value of textarea
message.value += ' Appended text.';

// --------------------------------------

//  get value of textarea on change
// message.addEventListener('input', function handleChange(event) {
//   console.log(event.target.value);
// });

function updateLabel() {
  label.textContent = message.value;
  var lines = message.value.split('\n');
  lines.forEach(element => {
    console.log(element);
  });
}

function onClick() {
  // if (inputsAreEmpty()) {
  //   label.textContent = 'Error: one or both inputs are empty.';
  //   return;
  // }

  updateLabel();
}

//var label = document.querySelector('p');
var label = document.getElementById('Label1');
var button = document.querySelector('button');
button.addEventListener('click', onClick);

