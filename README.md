officegen NodeRED Node
=====================

## Wrapper officegen
- https://www.npmjs.com/package/officegen
- Currently, only officegen's pptx generation is supported.
- plan to support other docs in the future. 

Install
-------

`npm install node-red-contrib-officegen`

Usage
-----

Expects a <b>msg.slides and msg.filename

## parameter example

```javascript
msg = {};
msg.filename = 'test.pptx';
msg.slides = [];
msg.slides[0] ={};
msg.slides[0].title = 'test ppt';
msg.slides[0].desc = 'test ppt desc';
msg.slides[0].name = 'test name';
// msg.slides[0].back = '000000';
msg.slides[0].adds =[];
msg.slides[0].adds[0] = {};
msg.slides[0].adds[0].addType = 'text';
msg.slides[0].adds[0].text = 'text message';
msg.slides[0].adds[0].textAttr = {
    y: 66,
    x: 'c',
    cx: '50%',
    cy: '1inch',
    font_size: 48,
    color: '0000ff',
    bodyProp: {
      normAutofit: 92500
    }
};

return msg;
```

## response
- result msg.payload : 
`completed` or error message

## sample flow

Flows can be imported and exported from the editor using their JSON format, making it very easy to share flows with others.

- Importing flows - pasting in the flow JSON directly
- Menu - Import - Clipboard - Ctrl+v - Import button 

```json
[{"id":"f48fa853.723fe8","type":"inject","z":"22dc8775.718778","name":"","props":[{"p":"payload"},{"p":"topic","vt":"str"}],"repeat":"","crontab":"","once":false,"onceDelay":0.1,"topic":"","payload":"","payloadType":"date","x":230,"y":80,"wires":[["861f37e3.d49be8"]]},{"id":"279d7707.53b9e8","type":"officegen","z":"22dc8775.718778","name":"make office file","x":560,"y":80,"wires":[["bc4de780.b165e8"]]},{"id":"861f37e3.d49be8","type":"function","z":"22dc8775.718778","name":"","func":"msg = {};\nmsg.filename = 'test.pptx';\nmsg.slides = [];\nmsg.slides[0] ={};\nmsg.slides[0].title = 'test ppt';\nmsg.slides[0].desc = 'test ppt desc';\nmsg.slides[0].name = 'test name';\n// msg.slides[0].back = '000000';\nmsg.slides[0].adds =[];\nmsg.slides[0].adds[0] = {};\nmsg.slides[0].adds[0].addType = 'text';\nmsg.slides[0].adds[0].text = 'text message';\nmsg.slides[0].adds[0].textAttr = {\n    y: 66,\n    x: 'c',\n    cx: '50%',\n    cy: '1inch',\n    font_size: 48,\n    color: '0000ff',\n    bodyProp: {\n      normAutofit: 92500\n    }\n};\n\nreturn msg;","outputs":1,"noerr":0,"initialize":"","finalize":"","x":380,"y":80,"wires":[["279d7707.53b9e8"]]},{"id":"bc4de780.b165e8","type":"debug","z":"22dc8775.718778","name":"","active":true,"tosidebar":true,"console":false,"tostatus":false,"complete":"false","statusVal":"","statusType":"auto","x":750,"y":80,"wires":[]}]
```
