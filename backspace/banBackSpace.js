//处理键盘事件 禁止后退键（Backspace）密码或单行、多行文本框除外 
function banBackSpace(e){
  var ev = e || window.event;
  //各种浏览器下获取事件对象
  var obj = ev.relatedTarget || ev.srcElement || ev.target ||ev.currentTarget;
  //按下Backspace键
  if(ev.keyCode == 8){
    var tagName = obj.nodeName //标签名称
    //如果标签不是input或者textarea则阻止Backspace
    if(tagName!='INPUT' && tagName!='TEXTAREA'){
      return stopIt(ev);
    }
    var tagType = obj.type.toUpperCase();//标签类型
    //input标签除了下面几种类型，全部阻止Backspace
    if(tagName=='INPUT' && (tagType!='TEXT' && tagType!='TEXTAREA' && tagType!='PASSWORD')){
      return stopIt(ev);
    }
    //input或者textarea输入框如果不可编辑则阻止Backspace
    if((tagName=='INPUT' || tagName=='TEXTAREA') && (obj.readOnly==true || obj.disabled ==true)){
      return stopIt(ev);
    }
  }
}

function stopIt(ev){
  if(ev.preventDefault ){
    //preventDefault()方法阻止元素发生默认的行为
    ev.preventDefault();
  }
  if(ev.returnValue){
    //IE浏览器下用window.event.returnValue = false;实现阻止元素发生默认的行为
    ev.returnValue = false;
  }
  return false;
}

//禁止后退键 作用于Firefox、Opera 
document.onkeypress = banBackSpace; 
//禁止后退键 作用于IE、Chrome 
document.onkeydown = banBackSpace;