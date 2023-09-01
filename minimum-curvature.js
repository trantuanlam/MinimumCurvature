const ZERO = 0.0000001;

class mcCoord {
  z; // TVD
  x; // North
  y; // East
/*  
  constructor (aX, aY, aZ) () {
    this.x = aX;
    this.y = aY;
    this.z = aZ;
  }
*/
  get neg() {
    let c = new mcCoord;
    c.x = -this.x;
    c.y = -this.y;
    c.z = -this.z;
    return c;
  }
  get negZ() {
    let c = new mcCoord;
    c.x = this.x;
    c.y = this.y;
    c.z = -this.z;
    return c;
  }

  addCoord(c1) {
    let c = new mcCoord;
    c.x = this.x + c1.x;
    c.y = this.y + c1.y;
    c.z = this.z + c1.z;
    return c;
  }

  addCoord2(c1, c2) {
    let c = new mcCoord;
    c.x = this.x + c1.x + c2.x;
    c.y = this.y + c1.y + c2.y;
    c.z = this.z + c1.z + c2.z;
    return c;
  }

  scale(s) {
    this.x = this.x * s;
    this.y = this.y * s;
    this.z = this.z * s;
  }

  translate(aX, aY, aZ) {
    this.x = this.x + aX;
    this.y = this.y + aY;
    this.z = this.z + aY;
  }

  get length(){
    return Math.sqrt(this.x**2 + this.y**2+ this.z**2);
  }

  get unitVector(){
    let c = new mcCoord;
    let l = length;
    c.x = this.x /l;
    c.y = this.y /l;
    c.z = this.z /l;
    return c;
  }

  vectorTo(aC) {
    let c = new mcCoord;
    c.x = aC.x - this.x;
    c.y = aC.y - this.y;
    c.z = aC.z - this.z;
    return c;
  }
  scalar(c) {
    return (this.x * c.x + this.y * c.y+ this.z * c.z);
  }
  isZero() {
    return (Math.abs(this.x) < ZERO) && (Math.abs(this.y) < ZERO) && (Math.abs(this.z) < ZERO)
  }
}


/////////////////////////////////////
function makeCoord(aX, aY, aZ){
  let c = new mcCoord;
  c.x = aX; 
  c.y = aY;
  c.z = aZ;
  return c;
}

function inclRadToVector(angle,  azimuth) {
  // angle, azimuth in radian
  let c = new mcCoord;
  c.z = Math.cos(angle);
  c.x = Math.sin(angle) * Math.cos(azimuth);
  c.y = Math.sin(angle) * Math.sin(azimuth);
  return c;
}

function inclDegToVector(angle,  azimuth) {
  // angle, azimuth in Degress
    return inclRadToVector(angle*Math.PI/180, azimuth*Math.PI/180);
}

function linerCombine(a, c1, b, c2) {
  // return vector a.c1 +  b.c2
  let c = new mcCoord;
  c.x = a * c1.x  + b * c2.x;
  c.y = a * c1.y  + b * c2.y;
  c.z = a * c1.z  + b * c2.z;
  return c;
}

function floatEqual(D1 , D2 ) {
  // so sanh hai so voi do chinh xac Zero
  return  Math.abs(D1 - D2) < ZERO;
}
 
function radiustoDLS30(r) {
  //Return (30 / R) * 180 / Math.PI
  return (30 / r) * 180 / Math.PI;
}

function dls30ToRadius(dls30) {  
  return (30 / (dls30 * Math.PI / 180));
}

function normAzimuthDeg(azimuth) {
  // chuan hoa gia tri Azimuth 0..360 degree
  // chua xu ly truong hop goc nho hon -360 hoac lon hon 720 degree
  if (azimuth > 360) { 
      return azimuth - 360
  } else if(azimuth < 0) {
      return azimuth + 360
  } else {
      return azimuth;
  }
}

function mcGetCoordDeg(I1, I2, A1, A2, DL) {
   //' I1, I2, A1, A2 tinh bang Degree
   let c = new mcCoord;

  if ((Math.abs(I1) < ZERO) && (Math.abs(I2) < ZERO)) {
        // thang dung
        c.z = DL;
        c.x = 0;
        c.y = 0;
  } else {
    I1 = I1 * Math.PI / 180;
    I2 = I2 * Math.PI / 180;
    A1 = A1 * Math.PI / 180;
    A2 = A2 * Math.PI / 180;

    if ((Math.abs(I1 - I2) < ZERO) && (Math.abs(A1 - A2) < ZERO)) {
      // on dinh goc
      
      let dL1 = DL * Math.sin((I1 + I2) / 2);
      
      c.z = DL * Math.cos((I1 + I2) / 2);
      c.x = dL1 * Math.cos((A1 + A2) / 2);
      c.y = dL1 * Math.sin((A1 + A2) / 2);
    } else {
      let sinI1 = Math.sin(I1);
      let sinI2 = Math.sin(I2);
      let cosI1 = Math.cos(I1);
      let cosI2 = Math.cos(I2);
      
      let sinA1 = Math.sin(A1);
      let sinA2 = Math.sin(A2);
      let cosA1 = Math.cos(A1);
      let cosA2 = Math.cos(A2);

      // aDogleg = dolegRad(I1, I2, A1, A2)
      let aDogleg = Math.acos(cosI1 * cosI2 + sinI1 * sinI2 * Math.cos(A2 - A1));
      
      let RF = 2 / (aDogleg) * Math.tan(aDogleg / 2);
      let tmp = DL / 2 * RF;
      
      c.z = tmp * (cosI1 + cosI2);
      c.x = tmp * (sinI1 * cosA1 + sinI2 * cosA2);
      c.y = tmp * (sinI1 * sinA1 + sinI2 * sinA2);
    }
  }
  return c;
}

function mcNorth( I1, I2, A1, A2, DL) {
  let c = mcGetCoordDeg(I1, I2, A1, A2, DL);
  return c.x;
}

function mcEast( I1, I2, A1, A2, DL) {
  let c = mcGetCoordDeg(I1, I2, A1, A2, DL);
  return c.y;
}

function mcVertival( I1, I2, A1, A2, DL) {
  let c = mcGetCoordDeg(I1, I2, A1, A2, DL);
  return c.z;
}

function dirAngleDeg(aNorth, aEast) {
    // Azimuth in Degrees
    let t =0;
    if (Math.abs(aEast) > ZERO) {
        t = Math.atan2(aNorth, aEast)*180/Math.PI;
        if (t < 0) {t = t + 360};
    } else {
      t = 0;
    }    
    return t;
}

function angleFromZXYDeg(tvd, north , east ) {
  // return inclination in Degrees
  let l = Math.sqrt(tvd**2 + north**2 + east**2);
  let aZ = tvd / l;
  return Math.acos(aZ) * Math.PI/180;
}

function azimuthFromZXY(tvd, north , east ) {
  let t =0;
  if ((Math.abs(north) < 0) && (Math.abs(east) < 0)) {
    t = 0;
  } else {
    t = Math.atan2(north, east);
  };

  if (t < 0) { t = t + 2* Math.PI};  // normalize zaimuth

  t = t * 180/ Math.PI;

  return t;
}
