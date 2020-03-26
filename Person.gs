(function(global){

  var Person = function(record){
    var _name = record[0];
    var _attendance1 = record[1];
    var _timeLate1 = record[2];
    var _timeEarly1 = record[3];
    var _targetRate1 = record[4];
    var _numServed1 = record[5];
    var _numApplying1 = record[6];
    var _numApplyingHome1 = record[7];
    var _breakTimePlanned1 = record[8];
    var _breakTimeActual1 = record[9];
    var _attendance2 = record[10];
    var _timeLate2 = record[11];
    var _timeEarly2 = record[12];
    var _targetRate2 = record[13];
    var _numServed2 = record[14];
    var _numApplying2 = record[15];
    var _numApplyingHome2 = record[16];
    var _breakTimePlanned2 = record[17];
    var _breakTimeActual2 = record[18];


    Object.defineProperties(this, {
      name: {
        get: function(){
          return _name;
        }
      }
    });
  };

  Person.prototype.greet = function(){

  };

  Person.prototype.log = function(){

  };

  global.Person = Person;

})(this);
