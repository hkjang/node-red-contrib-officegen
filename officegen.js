const officegen = require('officegen');
const fs = require('fs');
module.exports = function (RED) {
  function officegenFunction(config) {
    RED.nodes.createNode(this, config);
    var self = this;
    this.on('input', function(msg) {
      const pptx = officegen('pptx');
      for(var i in msg.slides){
        let slide = pptx.makeTitleSlide(msg.slides[i].title, msg.slides[i].desc);
        slide = pptx.makeNewSlide();
        slide.name = msg.slides[i].name;
        slide.back = msg.slides[i].back;
        for(var j in msg.slides[i].adds){
          if(msg.slides[i].adds[j].addType ==='text'){ // add text
            slide.addText(msg.slides[i].adds[j].text, msg.slides[i].adds[j].textAttr);
          }else if(msg.slides[i].adds[j].addType ==='shape'){ // add shape
            slide.addShape(pptx.shapes[msg.slides[i].adds[j].shapeType], msg.slides[i].adds[j].shapeAttr);
          }else if(msg.slides[i].adds[j].addType ==='image'){ // add image
            slide.addShape(path.resolve(__dirname, msg.slides[i].adds[j].imagePath), msg.slides[i].adds[j].imageAttr);
          }else if(msg.slides[i].adds[j].addType ==='table'){ // add table
            slide.addTable([msg.slides[i].adds[j].headerRow,msg.slides[i].adds[j].dataRows], msg.slides[i].adds[j].tableAttr);
          }else if(msg.slides[i].adds[j].addType ==='chart'){ // add chart
            slide.addChart(msg.slides[i].adds[j].chartAttr);
          }else{// just text
            slide.addText(msg.slides[i].adds[j].text);
          }
        }
      }
      var out = fs.createWriteStream(msg.filename);
      pptx.generate(out);
      out.on('error', function (err) {
        if(err){
          msg.payload = err;
          self.send(msg);
        }
      });
      out.on('close', function() {
        msg.payload = 'completed';
        self.send(msg);
      })
    });
  }
  RED.nodes.registerType('officegen', officegenFunction);
};
