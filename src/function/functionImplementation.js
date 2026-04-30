import luckysheetConfigsetting from '../controllers/luckysheetConfigsetting';
import {luckysheet_calcADPMM, luckysheet_getcelldata, luckysheet_getValue, luckysheet_parseData} from './func';
import {inverse} from './matrix_methods';
import {getluckysheetfile, getRangetxt, getSheetIndex} from '../methods/get';
import menuButton from '../controllers/menuButton';
import luckysheetSparkline from '../controllers/sparkline';
import formula from '../global/formula';
import func_methods from '../global/func_methods';
import editor from '../global/editor';
import {diff, isdatetime} from '../global/datecontroll';
import {error, isRealNull, isRealNum, valueIsError} from '../global/validate';
import {jfrefreshgrid, jfrefreshgridall} from '../global/refresh';
import {genarate, update} from '../global/format';
import {orderbydata} from '../global/sort';
import {datagridgrowth, getcellvalue} from '../global/getdata';
import {ABCatNum, chatatABC, getObjType, numFormat} from '../utils/util';
import Store from '../store';
import dayjs from 'dayjs';
import numeral from 'numeral';
import {
    askAIData,
    companyTargetData,
    companyTargetData10,
    companyTargetData11,
    companyTargetData12,
    excelToArray,
    excelToLuckyArray,
    getAirTable
} from '../demoData/getTargetData';
import {setcellvalue} from '../global/setdata';
import cityData from "../data/citydata";

//公式函数计算
const functionImplementation = {
  'SUM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array' && !func_methods.isDyadicArr(data)){
            return formula.error.v;
          }

          dataArr = dataArr.concat(func_methods.getDataArr(data, true));
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
        }
        else{
          if(!isRealNum(data)){
            if(getObjType(data) == 'boolean'){
              dataArr.push(data ? 1 : 0);
            }
            else{
              return formula.error.v;
            }
          }
          else{
            dataArr.push(data);
          }
        }
      }

      let sum = 0;

      for(let i = 0; i < dataArr.length; i++){
        if(valueIsError(dataArr[i])){
          return dataArr[i];
        }

        if(!isRealNum(dataArr[i])){
          continue;
        }

        sum = luckysheet_calcADPMM(sum, '+', dataArr[i]);// parseFloat(dataArr[i]);
      }

      return sum;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'AVERAGE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array'){
            if(!func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            dataArr = dataArr.concat(func_methods.getDataArr(data, true));
          }
          else{
            dataArr = dataArr.concat(data);
          }
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
        }
        else{
          dataArr.push(data);
        }
      }

      let sum = 0;
      let count = 0;

      for(let i = 0; i < dataArr.length; i++){
        if(valueIsError(dataArr[i])){
          return dataArr[i];
        }

        if(!isRealNum(dataArr[i])){
          return formula.error.v;
        }

        sum = luckysheet_calcADPMM(sum, '+', dataArr[i]);// parseFloat(dataArr[i]);
        count++;
      }

      if(count == 0){
        return formula.error.d;
      }

      return luckysheet_calcADPMM(sum, '/', count);// sum / count;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUNT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array'){
            if(!func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            dataArr = dataArr.concat(func_methods.getDataArr(data, true));
          }
          else{
            dataArr = dataArr.concat(data);
          }
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
        }
        else{
          if(getObjType(data) == 'boolean'){
            dataArr.push(data ? 1 : 0);
          }
          else{
            dataArr.push(data);
          }
        }
      }

      let count = 0;

      for(let i = 0; i < dataArr.length; i++){
        if(isRealNum(dataArr[i])){
          count++;
        }
      }

      return count;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUNTA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array'){
            if(!func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            dataArr = dataArr.concat(func_methods.getDataArr(data));
          }
          else{
            dataArr = dataArr.concat(data);
          }
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
        }
        else{
          dataArr.push(data);
        }
      }

      return dataArr.length;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MAX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array'){
            if(!func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            dataArr = dataArr.concat(func_methods.getDataArr(data, true));
          }
          else{
            dataArr = dataArr.concat(data);
          }
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
        }
        else{
          dataArr.push(data);
        }
      }

      let max = null;

      for(let i = 0; i < dataArr.length; i++){
        if(valueIsError(dataArr[i])){
          return dataArr[i];
        }

        if(!isRealNum(dataArr[i])){
          continue;
        }

        const num = parseFloat(dataArr[i]);
        if(max == null || num > max){
          max = num;
        }
      }

      return max == null ? 0 : max;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MIN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array'){
            if(!func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            dataArr = dataArr.concat(func_methods.getDataArr(data, true));
          }
          else{
            dataArr = dataArr.concat(data);
          }
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
        }
        else{
          dataArr.push(data);
        }
      }

      let min = null;

      for(let i = 0; i < dataArr.length; i++){
        if(valueIsError(dataArr[i])){
          return dataArr[i];
        }

        if(!isRealNum(dataArr[i])){
          continue;
        }

        const num = parseFloat(dataArr[i]);
        if(min == null || num < min){
          min = num;
        }
      }

      return min == null ? 0 : min;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'AGE_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let birthday = window.luckysheet_function.BIRTHDAY_BY_IDCARD.f(UUserCard);
      if(valueIsError(birthday)){
        return birthday;
      }

      birthday = dayjs(birthday);

      let currentDate = dayjs();
      if(arguments.length == 2){
        currentDate = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(currentDate)){
          return currentDate;
        }

        currentDate = dayjs(currentDate);
      }

      const age = currentDate.diff(birthday, 'years');

      if(age < 0 || isNaN(age)){
        return formula.error.v;
      }

      return age;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SEX_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]).toString();
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      // 确保身份证号长度足够且第17位是有效数字
      if (UUserCard.length < 17) {
        return formula.error.v;
      }

      const genderDigit = parseInt(UUserCard.charAt(16), 10);
      if (isNaN(genderDigit)) {
        return formula.error.v;
      }

      return genderDigit % 2 === 1 ? '男' : '女';
    }
    catch (e) {
      //计算错误检测
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'BIRTHDAY_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]).toString();
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let birthday = '';
      if (UUserCard.length == 15) {
        birthday = `19${  UUserCard.substring(6, 8)  }/${  UUserCard.substring(8, 10)  }/${  UUserCard.substring(10, 12)}`;
      }
      else if (UUserCard.length == 18) {
        birthday = `${UUserCard.substring(6, 10)  }/${  UUserCard.substring(10, 12)  }/${  UUserCard.substring(12, 14)}`;
      }
      else {
        return formula.error.v;
      }

      //生日格式
      let datetype = 0;
      if (arguments[1] != null) {
        datetype = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(datetype)){
          return datetype;
        }
      }

      if(!isRealNum(datetype)){
        return formula.error.v;
      }

      datetype = parseInt(datetype);

      if(datetype < 0 || datetype > 2){
        return formula.error.v;
      }

      if(datetype == 0){
        return birthday;
      }
      else if(datetype == 1){
        return dayjs(birthday).format('YYYY-MM-DD');
      }
      else if(datetype == 2){
        return dayjs(birthday).format('YYYY年M月D日');
      }

      return formula.error.v;
    }
    catch (e) {
      //计算错误检测
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'PROVINCE_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]).toString();
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let native = '未知';
      const provinceArray = formula.classlist.province;

      if (UUserCard.substring(0, 2) in provinceArray) {
        native = provinceArray[UUserCard.substring(0, 2)];
      }

      return native;
    }
    catch (e) {
      //计算错误检测
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'CITY_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]).toString();
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let dataNum = cityData.length,
        native = '未知';

      for (let i = 0; i < dataNum; i++) {
        if (UUserCard.substring(0, 6) == cityData[i].code) {
          native = cityData[i].title;
          break;
        }
      }

      return native;
    }
    catch (e) {
      //计算错误检测
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'STAR_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let birthday = window.luckysheet_function.BIRTHDAY_BY_IDCARD.f(UUserCard);
      if(valueIsError(birthday)){
        return birthday;
      }

      birthday = new Date(birthday);

      const month = birthday.getMonth();
      const day = birthday.getDate();

      // 星座边界日期配置 [星座名称, 月份, 日期]
      const zodiacBoundaries = [
        ['魔羯座', 0, 1],   // 1月1日 - 1月19日
        ['水瓶座', 0, 20],  // 1月20日 - 2月18日
        ['双鱼座', 1, 19],  // 2月19日 - 3月20日
        ['白羊座', 2, 21],  // 3月21日 - 4月19日
        ['金牛座', 3, 21],  // 4月20日 - 5月20日
        ['双子座', 4, 21],  // 5月21日 - 6月21日
        ['巨蟹座', 5, 22],  // 6月22日 - 7月22日
        ['狮子座', 6, 23],  // 7月23日 - 8月22日
        ['处女座', 7, 23],  // 8月23日 - 9月22日
        ['天秤座', 8, 23],  // 9月23日 - 10月23日
        ['天蝎座', 9, 23],  // 10月24日 - 11月21日
        ['射手座', 10, 22], // 11月22日 - 12月21日
        ['魔羯座', 11, 22]  // 12月22日 - 12月31日
      ];

      // 从后往前查找匹配的星座
      for (let i = zodiacBoundaries.length - 1; i >= 0; i--) {
        const [signName, signMonth, signDay] = zodiacBoundaries[i];

        // 比较月份和日期确定星座
        if (month > signMonth || (month === signMonth && day >= signDay)) {
          return signName;
        }
      }

      return '未找到匹配星座信息';
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ANIMAL_BY_IDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //身份证号
      const UUserCard = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(UUserCard)){
        return UUserCard;
      }

      if (!window.luckysheet_function.ISIDCARD.f(UUserCard)) {
        return formula.error.v;
      }

      let birthday = window.luckysheet_function.BIRTHDAY_BY_IDCARD.f(UUserCard);
      if(valueIsError(birthday)){
        return birthday;
      }

      birthday = new Date(birthday);

      // 生肖列表，按顺序排列（猪、鼠、牛、虎、兔、龙、蛇、马、羊、猴、鸡、狗）
      const zodiacAnimals = ['猪', '鼠', '牛', '虎', '兔', '龙', '蛇', '马', '羊', '猴', '鸡', '狗'];

      // 计算生肖索引：(年份 + 9) % 12
      const year = birthday.getFullYear();
      const index = (year + 9) % 12;

      // 验证索引有效性并返回对应生肖
      if (index >= 0 && index < zodiacAnimals.length) {
        return zodiacAnimals[index];
      } else {
        return '未找到匹配生肖信息';
      }
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISIDCARD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const idcard = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(idcard)){
        return idcard;
      }

      // 身份证号正则表达式：支持15位、18位或17位+X/x格式
      const idCardRegex = /^(\d{15}|\d{18}|\d{17}[Xx])$/;

      return idCardRegex.test(idcard);
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DM_TEXT_CUTWORD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      //任意需要分词的文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      //分词模式
      let datetype = 0;
      if (arguments[1] != null) {
        datetype = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(datetype)){
          return datetype;
        }
      }

      if(!isRealNum(datetype)){
        return formula.error.v;
      }

      datetype = parseInt(datetype);

      // 验证分词模式是否为有效值（0、1或2）
      if(datetype !== 0 && datetype !== 1 && datetype !== 2){
        return formula.error.v;
      }

      // 发送异步请求进行文本分词
      $.post('/dataqk/tu/api/cutword', {
        'text': text,
        'type': datetype
      }, function(data) {
        const d = [].concat(Store.flowdata);
        formula.execFunctionGroup(cell_r, cell_c, data);
        d[cell_r][cell_c] = {
          'v': data,
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      });

      return 'loading...';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DM_TEXT_TFIDF': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 获取需要分词的文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if (valueIsError(text)) {
        return text;
      }

      // 获取关键词个数，默认值为20
      let keywordCount = 20;
      if (arguments[1] != null) {
        keywordCount = func_methods.getFirstValue(arguments[1]);
        if (valueIsError(keywordCount)) {
          return keywordCount;
        }
      }

      // 验证关键词个数是否为有效数字
      if (!isRealNum(keywordCount)) {
        return formula.error.v;
      }

      keywordCount = parseInt(keywordCount);

      // 获取语料库类型，默认值为0
      let corpusType = 0;
      if (arguments[2] != null) {
        corpusType = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(corpusType)) {
          return corpusType;
        }
      }

      // 验证语料库类型是否为有效数字
      if (!isRealNum(corpusType)) {
        return formula.error.v;
      }

      corpusType = parseInt(corpusType);

      // 验证参数范围
      if (keywordCount < 0) {
        return formula.error.v;
      }

      if (![0, 1, 2].includes(corpusType)) {
        return formula.error.v;
      }

      // 发送TF-IDF计算请求
      $.post('/dataqk/tu/api/tfidf', {
        'text': text,
        'count': keywordCount,
        'set': corpusType
      })
      .done(function(data) {
        try {
          const d = editor.deepCopyFlowData(Store.flowdata);
          formula.execFunctionGroup(cell_r, cell_c, data);
          d[cell_r][cell_c] = {
            'v': data,
            'f': cell_fp
          };
          jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
        } catch (innerError) {
          console.error('处理TF-IDF结果时出错:', innerError);
        }
      })
      .fail(function(jqXHR, textStatus, errorThrown) {
        console.error('TF-IDF API请求失败:', textStatus, errorThrown);
        // 可以在这里添加用户友好的错误提示
      });

      return 'loading...';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DM_TEXT_TEXTRANK': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 获取需要分词的文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if (valueIsError(text)) {
        return text;
      }

      // 获取关键词个数，默认值为20
      let keywordCount = 20;
      if (arguments[1] != null) {
        keywordCount = func_methods.getFirstValue(arguments[1]);
        if (valueIsError(keywordCount)) {
          return keywordCount;
        }
      }

      // 验证关键词个数是否为有效数字
      if (!isRealNum(keywordCount)) {
        return formula.error.v;
      }

      keywordCount = parseInt(keywordCount);

      // 获取语料库类型，默认值为0
      let corpusType = 0;
      if (arguments[2] != null) {
        corpusType = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(corpusType)) {
          return corpusType;
        }
      }

      // 验证语料库类型是否为有效数字
      if (!isRealNum(corpusType)) {
        return formula.error.v;
      }

      corpusType = parseInt(corpusType);

      // 验证参数范围
      if (keywordCount < 0) {
        return formula.error.v;
      }

      if (![0, 1, 2].includes(corpusType)) {
        return formula.error.v;
      }

      // 发送TextRank计算请求（注意：这里应该是textrank API而不是tfidf）
      $.post('/dataqk/tu/api/textrank', {
        'text': text,
        'count': keywordCount,
        'set': corpusType
      })
      .done(function(data) {
        try {
          const d = editor.deepCopyFlowData(Store.flowdata);
          formula.execFunctionGroup(cell_r, cell_c, data);
          d[cell_r][cell_c] = {
            'v': data,
            'f': cell_fp
          };
          jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
        } catch (innerError) {
          console.error('处理TextRank结果时出错:', innerError);
        }
      })
      .fail(function(jqXHR, textStatus, errorThrown) {
        console.error('TextRank API请求失败:', textStatus, errorThrown);
        // 可以在这里添加用户友好的错误提示
      });

      return 'loading...';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_CLOSE': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 获取股票代码
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      // 获取并处理日期参数
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        // 检查数据类型是否为数组，如果是则返回错误
        if (objType === 'array') {
          return formula.error.v;
        }
        // 检查是否为对象且有startCell属性（单元格引用）
        else if (objType === 'object' && data_date.startCell != null) {
          // 验证日期格式是否正确
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        }
        // 其他情况直接使用数据
        else {
          date = data_date;
        }

        // 验证日期格式是否正确
        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        // 格式化日期为 YYYY-MM-DD
        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 获取复权类型，默认值为0（不复权）
      let adjustmentType = 0;
      if (arguments[2] != null) {
        adjustmentType = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(adjustmentType)) {
          return adjustmentType;
        }
      }

      // 验证复权类型是否为有效数字
      if (!isRealNum(adjustmentType)) {
        return formula.error.v;
      }

      adjustmentType = parseInt(adjustmentType);

      // 验证复权类型范围（0: 不复权, 1: 前复权, 2: 后复权）
      if (![0, 1, 2].includes(adjustmentType)) {
        return formula.error.v;
      }

      // 发送股票收盘价查询请求
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': adjustmentType,
        'type': '0'
      })
      .done(function(data) {
        try {
          const d = editor.deepCopyFlowData(Store.flowdata);
          let v = numFormat(data);
          if (v == null) {
            v = data;
          }
          formula.execFunctionGroup(cell_r, cell_c, v);
          d[cell_r][cell_c] = {
            'v': v,
            'f': cell_fp
          };
          jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
        } catch (innerError) {
          console.error('处理股票收盘价结果时出错:', innerError);
        }
      })
      .fail(function(jqXHR, textStatus, errorThrown) {
        console.error('股票信息API请求失败:', textStatus, errorThrown);
        // 可以在这里添加用户友好的错误提示
      });

      return 'loading...';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_OPEN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 获取股票代码
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      // 获取并处理日期参数
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        // 检查数据类型是否为数组，如果是则返回错误
        if (objType === 'array') {
          return formula.error.v;
        }
        // 检查是否为对象且有startCell属性（单元格引用）
        else if (objType === 'object' && data_date.startCell != null) {
          // 验证日期格式是否正确
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        }
        // 其他情况直接使用数据
        else {
          date = data_date;
        }

        // 验证日期格式是否正确
        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        // 格式化日期为 YYYY-MM-DD
        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 获取复权类型，默认值为0（不复权）
      let adjustmentType = 0;
      if (arguments[2] != null) {
        adjustmentType = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(adjustmentType)) {
          return adjustmentType;
        }
      }

      // 验证复权类型是否为有效数字
      if (!isRealNum(adjustmentType)) {
        return formula.error.v;
      }

      adjustmentType = parseInt(adjustmentType);

      // 验证复权类型范围（0: 不复权, 1: 前复权, 2: 后复权）
      if (![0, 1, 2].includes(adjustmentType)) {
        return formula.error.v;
      }

      // 发送股票开盘价查询请求（type: '1' 表示开盘价）
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': adjustmentType,
        'type': '1'
      })
      .done(function(data) {
        try {
          const d = editor.deepCopyFlowData(Store.flowdata);
          formula.execFunctionGroup(cell_r, cell_c, data);
          d[cell_r][cell_c] = {
            'v': data,
            'f': cell_fp
          };
          jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
        } catch (innerError) {
          console.error('处理股票开盘价结果时出错:', innerError);
        }
      })
      .fail(function(jqXHR, textStatus, errorThrown) {
        console.error('股票信息API请求失败:', textStatus, errorThrown);
        // 可以在这里添加用户友好的错误提示
      });

      return 'loading...';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_MAX': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 股票代码验证
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      if (!stockcode || typeof stockcode !== 'string') {
        return formula.error.v;
      }

      // 日期处理与验证
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        if (objType === 'array') {
          return formula.error.v;
        } else if (objType === 'object' && data_date.startCell != null) {
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        } else {
          date = data_date;
        }

        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 复权除权参数验证
      let price = 0;
      if (arguments[2] != null) {
        price = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(price)) {
          return price;
        }
      }

      if (!isRealNum(price)) {
        return formula.error.v;
      }

      price = parseInt(price, 10);

      if (![0, 1, 2].includes(price)) {
        return formula.error.v;
      }

      // 异步获取股票数据 - 优化错误处理和回调逻辑
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': price,
        type: '2'
      })
      .done(function(data) {
        const d = editor.deepCopyFlowData(Store.flowdata);
        formula.execFunctionGroup(cell_r, cell_c, data);
        d[cell_r][cell_c] = {
          'v': data,
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      })
      .fail(function(xhr, status, error) {
        console.error('获取股票数据失败:', status, error);
        // 在单元格中显示错误信息而不是静默失败
        const d = editor.deepCopyFlowData(Store.flowdata);
        d[cell_r][cell_c] = {
          'v': '获取数据失败',
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      });

      return 'loading...';
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_MIN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 股票代码验证
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      if (!stockcode || typeof stockcode !== 'string') {
        return formula.error.v;
      }

      // 日期处理与验证
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        if (objType === 'array') {
          return formula.error.v;
        } else if (objType === 'object' && data_date.startCell != null) {
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        } else {
          date = data_date;
        }

        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 复权除权参数验证
      let price = 0;
      if (arguments[2] != null) {
        price = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(price)) {
          return price;
        }
      }

      if (!isRealNum(price)) {
        return formula.error.v;
      }

      price = parseInt(price, 10);

      if (![0, 1, 2].includes(price)) {
        return formula.error.v;
      }

      // 异步获取股票数据 - 优化错误处理和回调逻辑
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': price,
        type: '3'
      })
      .done(function(data) {
        const d = editor.deepCopyFlowData(Store.flowdata);
        formula.execFunctionGroup(cell_r, cell_c, data);
        d[cell_r][cell_c] = {
          'v': data,
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      })
      .fail(function(xhr, status, error) {
        console.error('获取股票数据失败:', status, error);
        // 在单元格中显示错误信息而不是静默失败
        const d = editor.deepCopyFlowData(Store.flowdata);
        d[cell_r][cell_c] = {
          'v': '获取数据失败',
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      });

      return 'loading...';
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_VOLUMN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 股票代码验证
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      if (!stockcode || typeof stockcode !== 'string') {
        return formula.error.v;
      }

      // 日期处理与验证
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        if (objType === 'array') {
          return formula.error.v;
        } else if (objType === 'object' && data_date.startCell != null) {
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        } else {
          date = data_date;
        }

        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 复权除权参数验证
      let price = 0;
      if (arguments[2] != null) {
        price = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(price)) {
          return price;
        }
      }

      if (!isRealNum(price)) {
        return formula.error.v;
      }

      price = parseInt(price, 10);

      if (![0, 1, 2].includes(price)) {
        return formula.error.v;
      }

      // 异步获取股票数据 - 优化错误处理和回调逻辑
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': price,
        type: '4'
      })
      .done(function(data) {
        const d = editor.deepCopyFlowData(Store.flowdata);
        formula.execFunctionGroup(cell_r, cell_c, data);
        d[cell_r][cell_c] = {
          'v': data,
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      })
      .fail(function(xhr, status, error) {
        console.error('获取股票数据失败:', status, error);
        // 在单元格中显示错误信息而不是静默失败
        const d = editor.deepCopyFlowData(Store.flowdata);
        d[cell_r][cell_c] = {
          'v': '获取数据失败',
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      });

      return 'loading...';
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'DATA_CN_STOCK_AMOUNT': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;

      // 股票代码验证
      const stockcode = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(stockcode)) {
        return stockcode;
      }

      if (!stockcode || typeof stockcode !== 'string') {
        return formula.error.v;
      }

      // 日期处理与验证
      let date = null;
      if (arguments[1] != null) {
        const data_date = arguments[1];
        const objType = getObjType(data_date);

        if (objType === 'array') {
          return formula.error.v;
        } else if (objType === 'object' && data_date.startCell != null) {
          if (data_date.data != null &&
              getObjType(data_date.data) !== 'array' &&
              data_date.data.ct != null &&
              data_date.data.ct.t === 'd') {
            date = update('yyyy-mm-dd', data_date.data.v);
          } else {
            return formula.error.v;
          }
        } else {
          date = data_date;
        }

        if (!isdatetime(date)) {
          return [formula.error.v, '日期错误'];
        }

        date = dayjs(date).format('YYYY-MM-DD');
      }

      // 复权除权参数验证
      let price = 0;
      if (arguments[2] != null) {
        price = func_methods.getFirstValue(arguments[2]);
        if (valueIsError(price)) {
          return price;
        }
      }

      if (!isRealNum(price)) {
        return formula.error.v;
      }

      price = parseInt(price, 10);

      if (![0, 1, 2].includes(price)) {
        return formula.error.v;
      }

      // 异步获取股票数据 - 优化错误处理和回调逻辑
      $.post('/dataqk/tu/api/getstockinfo', {
        'stockCode': stockcode,
        'date': date,
        'price': price,
        type: '5'
      })
      .done(function(data) {
        const d = editor.deepCopyFlowData(Store.flowdata);
        formula.execFunctionGroup(cell_r, cell_c, data);
        d[cell_r][cell_c] = {
          'v': data,
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      })
      .fail(function(xhr, status, error) {
        console.error('获取股票数据失败:', status, error);
        // 在单元格中显示错误信息而不是静默失败
        const d = editor.deepCopyFlowData(Store.flowdata);
        d[cell_r][cell_c] = {
          'v': '获取数据失败',
          'f': cell_fp
        };
        jfrefreshgrid(d, [{'row': [cell_r, cell_r], 'column': [cell_c, cell_c]}]);
      });

      return 'loading...';
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ISDATE': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 日期验证
      const date = func_methods.getFirstValue(arguments[0], 'text');
      if (valueIsError(date)) {
        return date;
      }

      return isdatetime(date);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SUMIF': function() {
    // 必要参数个数错误检测（严格校验 2-3 个参数）
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //=SUMIF(A2:A5,">1600000",B2:B5)
      //=SUMIF(A2:A5,">1600000")
      //=SUMIF(A2:A5,3000000,B2:B5)
      let sum = 0;

      // 条件区域数据处理
      let rangeData = arguments[0].data;
      const rangeRow = arguments[0].rowl;
      const rangeCol = arguments[0].coll;
      // 解析条件
      const criteria = luckysheet_parseData(arguments[1]);
      // 转为一维数组，统一遍历逻辑
      rangeData = formula.getRangeArray(rangeData)[0];

      // 求和区域处理
      let sumRangeData = [];
      if (arguments[2]) {
        const sumRangeStart = arguments[2].startCell;
        const sumRangeRow = arguments[2].rowl;
        const sumRangeCol = arguments[2].coll;
        const sumRangeSheet = arguments[2].sheetName;

        // 尺寸一致直接使用，不一致则自动截取匹配尺寸的区域（Excel 标准规则）
        if (rangeRow === sumRangeRow && rangeCol === sumRangeCol) {
          sumRangeData = arguments[2].data;
        } else {
          // 解析起始单元格行号列号
          const startRow = parseInt(sumRangeStart.replace(/[^0-9]/g, '')) - 1;
          const startCol = ABCatNum(sumRangeStart.replace(/[^A-Za-z]/g, ''));
          // 计算结束位置（与条件区域同大小）
          const endRow = startRow + rangeRow - 1;
          const endCol = startCol + rangeCol - 1;
          // 转换为单元格地址
          const endABC = chatatABC(endCol);
          const endNum = endRow + 1;
          const sumRangeEnd = endABC + endNum;
          // 获取正确尺寸的求和区域数据
          const realSumRange = `${sumRangeSheet}!${sumRangeStart}:${sumRangeEnd}`;
          sumRangeData = luckysheet_getcelldata(realSumRange).data;
        }
        // 转为一维数组
        sumRangeData = formula.getRangeArray(sumRangeData)[0];
      }

      // 核心遍历：统一 2参数/3参数 逻辑，简化代码
      for (let i = 0; i < rangeData.length; i++) {
        const conditionVal = rangeData[i];
        // 条件匹配判断（修复：不再跳过 0、空值、布尔值）
        if (formula.acompareb(conditionVal, criteria)) {
          let addVal;
          // 区分有无求和区域
          if (arguments[2]) {
            addVal = sumRangeData[i];
          } else {
            addVal = conditionVal;
          }
          // 仅数字参与求和（兼容标准规则）
          if (isRealNum(addVal)) {
            // 高精度计算求和
            sum = luckysheet_calcADPMM(sum, '+', Number(addVal));
          }
        }
      }

      return sum;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'TAN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.tan(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'TANH': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      const e2 = Math.exp(2 * number);

      return (e2 - 1) / (e2 + 1);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'CEILING': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // number
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      // significance
      let significance = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(significance)) {
        return significance;
      }

      if (!isRealNum(significance)) {
        return formula.error.v;
      }

      significance = parseFloat(significance);

      if (significance == 0) {
        return 0;
      }

      if (number > 0 && significance < 0) {
        return formula.error.nm;
      }

      return Math.ceil(number / significance) * significance;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ATAN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.atan(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ASINH': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.log(number + Math.sqrt(number * number + 1));
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ABS': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.abs(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ACOS': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      if (number < -1 || number > 1) {
        return formula.error.nm;
      }

      return Math.acos(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ACOSH': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      if (number < 1) {
        return formula.error.nm;
      }

      return Math.log(number + Math.sqrt(number * number - 1));
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'MULTINOMIAL': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];

        if (getObjType(data) == 'array') {
          if (getObjType(data[0]) == 'array' && !func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }

          dataArr = dataArr.concat(func_methods.getDataArr(data, true));
        } else if (getObjType(data) == 'object' && data.startCell != null) {
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
        } else {
          dataArr.push(data);
        }
      }

      let sum = 0;
      let divisor = 1;

      for (let i = 0; i < dataArr.length; i++) {
        let number = dataArr[i];

        if (!isRealNum(number)) {
          return formula.error.v;
        }

        number = parseFloat(number);

        if (number < 0) {
          return formula.error.nm;
        }

        sum += number;
        divisor *= func_methods.factorial(number);
      }

      return func_methods.factorial(sum) / divisor;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ATANH': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      if (number <= -1 || number >= 1) {
        return formula.error.nm;
      }

      return Math.log((1 + number) / (1 - number)) / 2;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ATAN2': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 要计算其与x轴夹角大小的线段的终点x坐标
      let number_x = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number_x)) {
        return number_x;
      }

      if (!isRealNum(number_x)) {
        return formula.error.v;
      }

      number_x = parseFloat(number_x);

      // 要计算其与x轴夹角大小的线段的终点y坐标
      let number_y = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(number_y)) {
        return number_y;
      }

      if (!isRealNum(number_y)) {
        return formula.error.v;
      }

      number_y = parseFloat(number_y);

      if (number_x == 0 && number_y == 0) {
        return formula.error.d;
      }

      return Math.atan2(number_y, number_x);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'COUNTBLANK': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const data = arguments[0];
      let sum = 0;

      // 处理对象类型数据（带startCell的单元格区域）
      if (getObjType(data) === 'object' && data.startCell != null) {
        if (data.data == null) {
          return 1;
        }

        // 处理数组类型数据（二维数组表示的区域）
        if (getObjType(data.data) === 'array') {
          for (let r = 0; r < data.data.length; r++) {
            for (let c = 0; c < data.data[r].length; c++) {
              if (data.data[r][c] == null || isRealNull(data.data[r][c].v)) {
                sum++;
              }
            }
          }
        }
        // 处理单个单元格数据
        else {
          if (isRealNull(data.data.v)) {
            sum++;
          }
        }
      }

      return sum;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'COSH': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);
      return (Math.exp(number) + Math.exp(-number)) / 2;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'INT': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const data = arguments[0];

      // 处理数组类型数据
      if (getObjType(data) === 'array') {
        if (getObjType(data[0]) === 'array') {
          if (!func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }

          if (!isRealNum(data[0][0])) {
            return formula.error.v;
          }

          return Math.floor(parseFloat(data[0][0]));
        } else {
          if (!isRealNum(data[0])) {
            return formula.error.v;
          }

          return Math.floor(parseFloat(data[0]));
        }
      }
      // 处理对象类型数据（带startCell的单元格区域）
      else if (getObjType(data) === 'object' && data.startCell != null) {
        if (data.coll > 1) {
          return formula.error.v;
        }

        let cell;
        if (data.rowl > 1) {
          const cellrange = formula.getcellrange(data.startCell);
          const str = cellrange.row[0];

          if (window.luckysheetCurrentRow < str || window.luckysheetCurrentRow > str + data.rowl - 1) {
            return formula.error.v;
          }

          cell = data.data[window.luckysheetCurrentRow - str][0];
        } else {
          cell = data.data;
        }

        if (cell == null || isRealNull(cell.v)) {
          return 0;
        }

        if (!isRealNum(cell.v)) {
          return formula.error.v;
        }

        return Math.floor(parseFloat(cell.v));
      }
      // 处理其他类型数据（包括布尔值）
      else {
        if (getObjType(data) === 'boolean') {
          if (data.toString().toLowerCase() === 'true') {
            return 1;
          }

          if (data.toString().toLowerCase() === 'false') {
            return 0;
          }
        }

        if (!isRealNum(data)) {
          return formula.error.v;
        }

        return Math.floor(parseFloat(data));
      }
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ISEVEN': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseInt(number);
      return Math.abs(number) & 1 ? false : true;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ISODD': function() {
    // 必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseInt(number);
      return Math.abs(number) & 1 ? true : false;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  //1111
  'LCM': function() {
    // 参数个数校验（最少1个，最多255个，Excel标准）
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let numberList = [];

      // 统一收集所有参数（支持数字、一维数组、单元格区域）
      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];
        const dataType = getObjType(data);

        if (dataType === 'array') {
          // 二维数组判断
          if (!func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }
          // 提取数组数据
          numberList = numberList.concat(func_methods.getDataArr(data));
        } else if (dataType === 'object' && data.startCell != null) {
          // 单元格区域提取数字
          numberList = numberList.concat(func_methods.getCellDataArr(data, 'number', true));
        } else {
          // 普通数字
          numberList.push(data);
        }
      }

      // 数据清洗：验证数字 + 转为正整数（Excel标准）
      const validNumbers = [];
      for (let num of numberList) {
        // 非数字直接返回错误
        if (!isRealNum(num)) {
          return formula.error.v;
        }

        // 转为整数
        const integer = Math.floor(Number(num));

        // 负数返回 #NUM!
        if (integer < 0) {
          return formula.error.nm;
        }

        validNumbers.push(integer);
      }

      // 包含 0 直接返回 0（Excel标准规则）
      if (validNumbers.includes(0)) {
        return 0;
      }

      // 空数据返回 0
      if (validNumbers.length === 0) {
        return 0;
      }

      // ------------------------------
      // 最小公倍数核心算法（清晰版）
      // ------------------------------
      // 最大公约数
      const gcd = (a, b) => {
        while (b !== 0) {
          let temp = b;
          b = a % b;
          a = temp;
        }
        return a;
      };

      // 两个数的最小公倍数
      const lcmTwo = (a, b) => {
        if (a === 0 || b === 0) return 0;
        return (a * b) / gcd(a, b);
      };

      // 多个数的最小公倍数
      let result = validNumbers[0];
      for (let i = 1; i < validNumbers.length; i++) {
        result = lcmTwo(result, validNumbers[i]);

        // 超出 JS 安全数字范围返回错误
        if (result >= Math.pow(2, 53)) {
          return formula.error.nm;
        }
      }

      return result;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'LN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 获取参数第一个值（兼容区域/数组）
      let number = func_methods.getFirstValue(arguments[0]);

      // 如果是错误值，直接返回错误
      if (valueIsError(number)) {
        return number;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      // 转换为浮点数字
      number = parseFloat(number);

      // 小于等于 0 返回 #NUM!（Excel 标准规则）
      if (number <= 0) {
        return formula.error.nm;
      }

      // 计算自然对数
      return Math.log(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'LOG': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 提取并校验第一个参数：待计算数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 必须大于 0
      if (number <= 0) {
        return formula.error.nm;
      }

      // 提取并校验第二个参数：底数（默认 10）
      let base = 10;
      if (arguments.length === 2) {
        base = func_methods.getFirstValue(arguments[1]);
        if (valueIsError(base)) {
          return base;
        }
        if (!isRealNum(base)) {
          return formula.error.v;
        }
        base = parseFloat(base);

        // 底数必须 > 0 且 ≠ 1
        if (base <= 0 || base === 1) {
          return formula.error.nm;
        }
      }

      // 计算对数：换底公式
      return Math.log(number) / Math.log(base);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'LOG10': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 获取参数第一个有效值
      let number = func_methods.getFirstValue(arguments[0]);

      // 透传错误值
      if (valueIsError(number)) {
        return number;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      // 转换为数字
      number = parseFloat(number);

      // 小于等于 0 返回 #NUM!
      if (number <= 0) {
        return formula.error.nm;
      }

      // 计算以 10 为底的对数（可直接用 Math.log10，更标准）
      return Math.log10(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'MOD': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 被除数
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 除数
      let divisor = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(divisor)) {
        return divisor;
      }
      if (!isRealNum(divisor)) {
        return formula.error.v;
      }
      divisor = parseFloat(divisor);

      // 除数不能为 0，返回 #DIV/0!
      if (divisor === 0) {
        return formula.error.d;
      }

      // Excel 标准 MOD 算法（结果符号与除数一致）
      const mod = number - (divisor * Math.floor(number / divisor));

      return mod;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'MROUND': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 获取要舍入的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 获取要舍入到的倍数
      let multiple = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(multiple)) {
        return multiple;
      }
      if (!isRealNum(multiple)) {
        return formula.error.v;
      }
      multiple = parseFloat(multiple);

      // Excel 标准规则：倍数不能为 0
      if (multiple === 0) {
        return formula.error.nm;
      }

      // Excel 标准规则：数值与倍数必须同正负号
      if (Math.sign(number) !== Math.sign(multiple)) {
        return formula.error.nm;
      }

      // 核心计算：四舍五入到指定倍数
      return Math.round(number / multiple) * multiple;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ODD': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 获取参数值
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // Excel 标准 ODD 算法
      if (number >= 0) {
        // 正数：向上取奇数
        return Math.ceil(number) | 1;
      } else {
        // 负数：向下取奇数
        return Math.floor(number) | 1;
      }
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SUMSQ': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      // 遍历所有参数，收集数字（支持数组、单元格、普通数值）
      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];
        const dataType = getObjType(data);

        if (dataType === 'array') {
          // 非法二维数组直接返回错误
          if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }
          dataArr = dataArr.concat(func_methods.getDataArr(data, true));
        }
        else if (dataType === 'object' && data.startCell != null) {
          // 单元格区域提取数字
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
        }
        else {
          // 普通数值直接加入
          dataArr.push(data);
        }
      }

      // 平方求和（非数字直接返回 #VALUE!）
      let sum = 0;
      for (let num of dataArr) {
        if (!isRealNum(num)) {
          return formula.error.v;
        }
        const number = parseFloat(num);
        sum += number * number;
      }

      return sum;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'COMBIN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 总项目数
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = Math.floor(Number(number));

      // 选取项目数
      let number_chosen = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(number_chosen)) {
        return number_chosen;
      }
      if (!isRealNum(number_chosen)) {
        return formula.error.v;
      }
      number_chosen = Math.floor(Number(number_chosen));

      // Excel 标准错误规则
      if (number < 0 || number_chosen < 0 || number < number_chosen) {
        return formula.error.nm;
      }

      // 边界值：选取数为 0 直接返回 1
      if (number_chosen === 0) {
        return 1;
      }

      // 计算组合数（公式：C(n,k) = n!/(k!*(n-k)!)
      return func_methods.factorial(number) /
          (func_methods.factorial(number_chosen) * func_methods.factorial(number - number_chosen));
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SUBTOTAL': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 提取第一个参数：function_num
      const funcNumInput = arguments[0];
      let functionNum;

      // 获取 function_num 的真实值（支持区域/数组）
      if (getObjType(funcNumInput) === 'object' && funcNumInput.startCell != null) {
        functionNum = func_methods.getFirstValue(funcNumInput);
      } else if (getObjType(funcNumInput) === 'array') {
        functionNum = func_methods.getDataArr(funcNumInput);
      } else {
        functionNum = funcNumInput;
      }

      // 截取剩余参数
      const args = Array.from(arguments).slice(1);

      // 核心计算函数
      const compute = (num) => {
        num = parseInt(num);

        // Excel 标准：仅支持 1-11 / 101-111
        if (num < 1 || num > 111 || (num > 11 && num < 101)) {
          return formula.error.v;
        }

        const f = window.luckysheet_function;
        switch (num) {
          case 1: case 101: return f.AVERAGE.f.apply(f.AVERAGE, args);
          case 2: case 102: return f.COUNT.f.apply(f.COUNT, args);
          case 3: case 103: return f.COUNTA.f.apply(f.COUNTA, args);
          case 4: case 104: return f.MAX.f.apply(f.MAX, args);
          case 5: case 105: return f.MIN.f.apply(f.MIN, args);
          case 6: case 106: return f.PRODUCT.f.apply(f.PRODUCT, args);
          case 7: case 107: return f.STDEVA.f.apply(f.STDEVA, args);
          case 8: case 108: return f.STDEVP.f.apply(f.STDEVP, args);
          case 9: case 109: return f.SUM.f.apply(f.SUM, args);
          case 10: case 110: return f.VAR_S.f.apply(f.VAR_S, args);
          case 11: case 111: return f.VAR_P.f.apply(f.VAR_P, args);
          default: return formula.error.v;
        }
      };

      // 处理数组/单元格区域类型的 function_num
      if (getObjType(functionNum) === 'array') {
        const result = [];
        const is2D = getObjType(functionNum[0]) === 'array';

        for (let i = 0; i < functionNum.length; i++) {
          const row = is2D ? functionNum[i] : [functionNum[i]];
          const resRow = [];

          for (const val of row) {
            if (valueIsError(val)) {
              resRow.push(val);
            } else if (!isRealNum(val)) {
              resRow.push(formula.error.v);
            } else {
              resRow.push(compute(val));
            }
          }

          result.push(is2D ? resRow : resRow[0]);
        }
        return result;
      }

      // 处理单个值
      if (valueIsError(functionNum)) return functionNum;
      if (!isRealNum(functionNum)) return formula.error.v;

      return compute(functionNum);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ASIN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 获取参数值
      let number = func_methods.getFirstValue(arguments[0]);

      // 错误值直接返回
      if (valueIsError(number)) {
        return number;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      // 转换为浮点数
      number = parseFloat(number);

      // 值域必须在 [-1, 1] 之间，否则返回 #NUM!
      if (number < -1 || number > 1) {
        return formula.error.nm;
      }

      // 计算反正弦值（弧度）
      return Math.asin(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'COUNTIF': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 1. 获取条件区域（必须是单元格区域）
      const rangeArg = arguments[0];
      if (getObjType(rangeArg) !== 'object' || rangeArg.startCell == null) {
        return formula.error.v;
      }
      const rangeData = rangeArg.data;

      // 2. 获取条件
      const criteriaArg = arguments[1];
      let criteria;

      if (getObjType(criteriaArg) === 'array') {
        criteria = criteriaArg;
      } else if (getObjType(criteriaArg) === 'object' && criteriaArg.startCell != null) {
        // 条件为区域时，仅取左上角第一个值
        criteria = criteriaArg.data && criteriaArg.data[0] ? criteriaArg.data[0].v : '';
      } else {
        criteria = luckysheet_parseData(criteriaArg);
      }

      // 3. 核心匹配计算函数
      function getCriteriaResult(range, condition) {
        let cond = String(condition ?? '');

        // 标准化条件：无操作符则默认等于
        if (!/[<>=!*?]/.test(cond)) {
          cond = `==${cond}`;
        }
        // 替换不等号
        cond = cond.replace('<>', '!=');

        let count = 0;

        // 遍历区域
        if (Array.isArray(range)) {
          for (let row of range) {
            if (!Array.isArray(row)) row = [row];
            for (let cell of row) {
              if (cell == null || isRealNull(cell.v)) continue;
              const val = cell.v;

              // 通配符匹配
              if (cond.includes('*') || cond.includes('?')) {
                if (formula.isWildcard(val, cond)) count++;
              } else {
                // 安全条件判断
                if (formula.acompareb(val, cond)) count++;
              }
            }
          }
        }
        return count;
      }

      // 4. 数组条件返回数组结果
      if (Array.isArray(criteria)) {
        const result = [];
        const is2D = Array.isArray(criteria[0]);

        for (let row of criteria) {
          const resRow = [];
          const items = is2D ? row : [row];
          for (let c of items) {
            resRow.push(getCriteriaResult(rangeData, c));
          }
          result.push(is2D ? resRow : resRow[0]);
        }
        return result;
      }

      // 单个条件返回数字
      return getCriteriaResult(rangeData, criteria);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'RADIANS': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return number * Math.PI / 180;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'RAND': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    try {
      // 生成 [0,1) 随机数，完全符合 Excel 标准
      return Math.random();
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'COUNTUNIQUE': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      // 收集所有参数数据
      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];
        const dataType = getObjType(data);

        if (dataType === 'array') {
          if (data.length === 0) continue;
          if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }
          dataArr = dataArr.concat(func_methods.getDataArr(data, true));
        }
        else if (dataType === 'object' && data.startCell != null) {
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'all', true));
        }
        else {
          dataArr.push(data);
        }
      }

      // 获取去重后的数组并返回长度
      const uniqueArr = window.luckysheet_function.UNIQUE.f(dataArr);
      return Array.isArray(uniqueArr) ? uniqueArr.length : 0;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'DEGREES': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return number * 180 / Math.PI;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ERFC': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return jStat.erfc(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'EVEN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      // 错误值直接返回
      if (valueIsError(number)) {
        return number;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      // Excel 标准 EVEN 算法
      if (number >= 0) {
        return Math.ceil(number) % 2 === 0 ? Math.ceil(number) : Math.ceil(number) + 1;
      } else {
        return Math.floor(number) % 2 === 0 ? Math.floor(number) : Math.floor(number) - 1;
      }
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'EXP': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.exp(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'FACT': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      // 错误值直接返回
      if (valueIsError(number)) {
        return number;
      }

      // 兼容布尔值：TRUE=1，FALSE=0
      if (getObjType(number) === 'boolean') {
        number = number ? 1 : 0;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      // 转换为整数
      number = parseInt(number);

      // 负数返回 #NUM!
      if (number < 0) {
        return formula.error.nm;
      }

      // 计算阶乘
      return func_methods.factorial(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'FACTDOUBLE': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      // 错误值直接返回
      if (valueIsError(number)) {
        return number;
      }

      // 兼容布尔值：TRUE=1，FALSE=0
      if (getObjType(number) === 'boolean') {
        number = number ? 1 : 0;
      }

      // 非数字返回 #VALUE!
      if (!isRealNum(number)) {
        return formula.error.v;
      }

      // 转换为整数
      number = parseInt(number);

      // 负数返回 #NUM!
      if (number < 0) {
        return formula.error.nm;
      }

      // 计算双倍阶乘
      return func_methods.factorialDouble(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'PI': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    try {
      return Math.PI;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'FLOOR': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 要舍入的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 舍入倍数
      let significance = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(significance)) {
        return significance;
      }
      if (!isRealNum(significance)) {
        return formula.error.v;
      }
      significance = parseFloat(significance);

      // Excel 标准错误规则
      if (significance === 0) {
        return formula.error.d;
      }
      if (number > 0 && significance < 0) {
        return formula.error.nm;
      }

      // Excel 标准 FLOOR 算法（无精度丢失）
      return Math.floor(number / significance) * significance;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'GCD': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let dataArr = [];

      // 收集所有参数
      for (let i = 0; i < arguments.length; i++) {
        const data = arguments[i];
        const dataType = getObjType(data);

        if (dataType === 'array') {
          if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
            return formula.error.v;
          }
          dataArr = dataArr.concat(func_methods.getDataArr(data, false));
        }
        else if (dataType === 'object' && data.startCell != null) {
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', false));
        }
        else {
          dataArr.push(data);
        }
      }

      // 最大安全整数限制
      const MAX_SAFE_INTEGER = Math.pow(2, 53);

      // 遍历计算最大公约数
      let result = 0;
      for (const val of dataArr) {
        if (!isRealNum(val)) {
          return formula.error.v;
        }

        let num = parseInt(val);
        if (num < 0 || num >= MAX_SAFE_INTEGER) {
          return formula.error.nm;
        }

        // 欧几里得算法
        while (num !== 0) {
          const temp = num;
          num = result % num;
          result = temp;
        }
      }

      return result;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'RANDBETWEEN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 最小值
      let bottom = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(bottom)) {
        return bottom;
      }
      if (!isRealNum(bottom)) {
        return formula.error.v;
      }
      bottom = parseInt(bottom);

      // 最大值
      let top = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(top)) {
        return top;
      }
      if (!isRealNum(top)) {
        return formula.error.v;
      }
      top = parseInt(top);

      // 错误判断：最小值大于最大值
      if (bottom > top) {
        return formula.error.nm;
      }

      // 生成 [bottom, top] 区间随机整数（Excel标准）
      return Math.floor(Math.random() * (top - bottom + 1)) + bottom;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ROUND': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 要四舍五入的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 四舍五入的位数
      let digits = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(digits)) {
        return digits;
      }
      if (!isRealNum(digits)) {
        return formula.error.v;
      }
      digits = parseInt(digits);

      // 标准四舍五入计算（兼容正负）
      const power = Math.pow(10, digits);
      return Math.round(number * power) / power;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ROUNDDOWN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 要向下舍入的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 舍入位数
      let digits = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(digits)) {
        return digits;
      }
      if (!isRealNum(digits)) {
        return formula.error.v;
      }
      digits = parseInt(digits);

      // Excel 标准 ROUNDDOWN 算法
      const power = Math.pow(10, digits);
      return (number >= 0 ? Math.floor(number * power) : Math.ceil(number * power)) / power;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'ROUNDUP': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 要向上舍入的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(number)) {
        return number;
      }
      if (!isRealNum(number)) {
        return formula.error.v;
      }
      number = parseFloat(number);

      // 舍入位数
      let digits = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(digits)) {
        return digits;
      }
      if (!isRealNum(digits)) {
        return formula.error.v;
      }
      digits = parseInt(digits);

      // Excel 标准 ROUNDUP 算法
      const power = Math.pow(10, digits);
      return (number >= 0 ? Math.ceil(number * power) : Math.floor(number * power)) / power;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SERIESSUM': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      // 幂级数的输入值
      let x = func_methods.getFirstValue(arguments[0]);
      if (valueIsError(x)) return x;
      if (!isRealNum(x)) return formula.error.v;
      x = parseFloat(x);

      // x 的首项乘幂
      let n = func_methods.getFirstValue(arguments[1]);
      if (valueIsError(n)) return n;
      if (!isRealNum(n)) return formula.error.v;
      n = parseFloat(n);

      // 级数中每一项的乘幂 n 的步长增加值
      let m = func_methods.getFirstValue(arguments[2]);
      if (valueIsError(m)) return m;
      if (!isRealNum(m)) return formula.error.v;
      m = parseFloat(m);

      // 与 x 的每个连续乘幂相乘的一组系数
      const dataCoefficients = arguments[3];
      let coefficients = [];
      const dataType = getObjType(dataCoefficients);

      if (dataType === 'array') {
        if (getObjType(dataCoefficients[0]) === 'array' && !func_methods.isDyadicArr(dataCoefficients)) {
          return formula.error.v;
        }
        coefficients = coefficients.concat(func_methods.getDataArr(dataCoefficients, false));
      } else if (dataType === 'object' && dataCoefficients.startCell != null) {
        coefficients = coefficients.concat(func_methods.getCellDataArr(dataCoefficients, 'number', false));
      } else {
        coefficients.push(dataCoefficients);
      }

      // 计算
      if (!isRealNum(coefficients[0])) return formula.error.v;
      let result = parseFloat(coefficients[0]) * Math.pow(x, n);

      for (let i = 1; i < coefficients.length; i++) {
        let number = coefficients[i];
        if (!isRealNum(number)) return formula.error.v;
        number = parseFloat(number);
        result += number * Math.pow(x, n + i * m);
      }

      return result;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SIGN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      // Excel 标准符号函数
      if (number > 0) return 1;
      if (number < 0) return -1;
      return 0;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SIN': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return Math.sin(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SINH': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      return (Math.exp(number) - Math.exp(-number)) / 2;
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
  'SQRT': function() {
    // 参数个数校验
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    // 参数类型校验
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);
      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let number = func_methods.getFirstValue(arguments[0]);

      if (valueIsError(number)) {
        return number;
      }

      if (!isRealNum(number)) {
        return formula.error.v;
      }

      number = parseFloat(number);

      // 负数返回 #NUM!
      if (number < 0) {
        return formula.error.nm;
      }

      return Math.sqrt(number);
    } catch (e) {
      const err = formula.errorInfo(e);
      return [formula.error.v, err];
    }
  },
    'SQRTPI': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let number = func_methods.getFirstValue(arguments[0]);

            if (valueIsError(number)) {
                return number;
            }

            if (!isRealNum(number)) {
                return formula.error.v;
            }

            number = parseFloat(number);

            // 负数返回错误 #NUM!
            if (number < 0) {
                return formula.error.nm;
            }

            return Math.sqrt(number * Math.PI);
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'GAMMALN': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let number = func_methods.getFirstValue(arguments[0]);

            if (valueIsError(number)) {
                return number;
            }

            if (!isRealNum(number)) {
                return formula.error.v;
            }

            number = parseFloat(number);

            // 小于等于 0 返回 #NUM!
            if (number <= 0) {
                return formula.error.nm;
            }

            return jStat.gammaln(number);
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'COS': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let number = func_methods.getFirstValue(arguments[0]);

            if (valueIsError(number)) {
                return number;
            }

            if (!isRealNum(number)) {
                return formula.error.v;
            }

            number = parseFloat(number);

            return Math.cos(number);
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'TRUNC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要截取的数字
            let number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number)) {
                return number;
            }
            if (!isRealNum(number)) {
                return formula.error.v;
            }
            number = parseFloat(number);

            // 截取位数（默认 0）
            let digits = 0;
            if (arguments.length >= 2) {
                digits = func_methods.getFirstValue(arguments[1]);
                if (valueIsError(digits)) return digits;
                if (!isRealNum(digits)) return formula.error.v;
                digits = parseInt(digits);
            }

            // 标准 TRUNC 截断算法
            const power = Math.pow(10, digits);
            return Math.trunc(number * power) / power;
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'QUOTIENT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 被除数
            let numerator = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(numerator)) {
                return numerator;
            }
            if (!isRealNum(numerator)) {
                return formula.error.v;
            }
            numerator = parseFloat(numerator);

            // 除数
            let denominator = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(denominator)) {
                return denominator;
            }
            if (!isRealNum(denominator)) {
                return formula.error.v;
            }
            denominator = parseFloat(denominator);

            // 除零错误
            if (denominator === 0) {
                return formula.error.d;
            }

            // 标准取整（直接截断小数）
            return Math.trunc(numerator / denominator);
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'POWER': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 底数
            let number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number)) {
                return number;
            }
            if (!isRealNum(number)) {
                return formula.error.v;
            }
            number = parseFloat(number);

            // 指数
            let power = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(power)) {
                return power;
            }
            if (!isRealNum(power)) {
                return formula.error.v;
            }
            power = parseFloat(power);

            // Excel 错误规则
            if (number === 0 && power === 0) {
                return formula.error.nm;
            }
            if (number < 0 && !Number.isInteger(power)) {
                return formula.error.nm;
            }

            return Math.pow(number, power);
        } catch (e) {
            const err = formula.errorInfo(e);
            return [formula.error.v, err];
        }
    },
    'SUMIFS': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let sum = 0;
            const args = arguments;
            luckysheet_getValue(args);

            // 求和区域
            const sumRange = formula.getRangeArray(args[0])[0];
            const rowCount = sumRange.length;
            const conditionResults = new Array(rowCount).fill(true);

            // 遍历多组条件区域与条件
            for (let i = 1; i < args.length; i += 2) {
                const criteriaRange = formula.getRangeArray(args[i])[0];
                const criteria = args[i + 1];

                // 逐行判断条件
                for (let j = 0; j < criteriaRange.length; j++) {
                    const value = criteriaRange[j];
                    conditionResults[j] = conditionResults[j] && !!value && formula.acompareb(value, criteria);
                }
            }

            // 符合条件求和
            for (let i = 0; i < rowCount; i++) {
                if (conditionResults[i]) {
                    sum = luckysheet_calcADPMM(sum, '+', sumRange[i]);
                }
            }

            return sum;
        } catch (e) {
            const err = e;
            return formula.errorInfo(err);
        }
    },
    'GET_TARGET': function() {
        try {
            const luckysheetCurrentIndex = window.luckysheetCurrentIndex;
            const currentSheetIndex = Store.currentSheetIndex;

            // 页面不匹配直接返回
            if (luckysheetCurrentIndex !== currentSheetIndex) {
                return;
            }

            const startRow = window.luckysheetCurrentRow;
            const startColumn = window.luckysheetCurrentColumn;
            const cell_fp = window.luckysheetCurrentFunction;

            setTimeout(() => {
                let d = editor.deepCopyFlowData(Store.flowdata);
                const target = excelToLuckyArray(companyTargetData);

                const targetRows = target.length;
                const targetCols = target[0]?.length || 0;
                const rowHeight = startRow + targetRows;
                const colWidth = startColumn + targetCols;

                // 自动扩展表格
                if (rowHeight >= d.length && colWidth >= d[0].length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, colWidth - d[0].length + 1);
                } else if (rowHeight >= d.length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, 0);
                } else if (colWidth >= d[0].length) {
                    d = datagridgrowth(d, 0, colWidth - d[0].length + 1);
                }

                // 填充目标数据
                target.forEach((row, r) => {
                    row.forEach((cell, c) => {
                        setcellvalue(startRow + r, startColumn + c, d, cell);
                    });
                });

                // 恢复公式
                d[startRow][startColumn].f = cell_fp;
                delete d[startRow][startColumn].m;

                // 刷新表格
                if (currentSheetIndex === Store.currentSheetIndex) {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    const file = Store.luckysheetfile[sheetIndex];

                    file.row = d.length;
                    file.data = d;

                    jfrefreshgridall(
                        d[0].length,
                        d.length,
                        d,
                        null,
                        Store.luckysheet_select_save,
                        'datachangeAll',
                        undefined,
                        undefined
                    );
                } else {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    Store.luckysheetfile[sheetIndex].data = d;
                }

            }, 300);

            return 'loading...';
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'GET_AIRTABLE_DATA': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const luckysheetCurrentIndex = window.luckysheetCurrentIndex;
            const currentSheetIndex = Store.currentSheetIndex;

            // 工作表不匹配直接返回
            if (luckysheetCurrentIndex !== currentSheetIndex) {
                return;
            }

            const startRow = window.luckysheetCurrentRow;
            const startColumn = window.luckysheetCurrentColumn;

            // 解析参数
            const url = func_methods.getFirstValue(arguments[0]);
            const sortIndex = func_methods.getFirstValue(arguments[1]);
            const sortOrder = func_methods.getFirstValue(arguments[2]);
            const cellFp = window.luckysheetCurrentFunction;

            let d = editor.deepCopyFlowData(Store.flowdata);

            // 获取 Airtable 数据并渲染
            getAirTable(url, sortIndex, sortOrder, (data) => {
                const targetRows = data.length;
                const targetCols = data[0]?.length || 0;
                const rowHeight = startRow + targetRows;
                const colWidth = startColumn + targetCols;

                // 自动扩展表格
                if (rowHeight >= d.length && colWidth >= d[0].length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, colWidth - d[0].length + 1);
                } else if (rowHeight >= d.length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, 0);
                } else if (colWidth >= d[0].length) {
                    d = datagridgrowth(d, 0, colWidth - d[0].length + 1);
                }

                // 填充数据
                data.forEach((row, r) => {
                    row.forEach((cell, c) => {
                        setcellvalue(startRow + r, startColumn + c, d, cell);
                    });
                });

                // 恢复公式
                d[startRow][startColumn].f = cellFp;
                delete d[startRow][startColumn].m;

                // 刷新表格
                if (currentSheetIndex === Store.currentSheetIndex) {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    const file = Store.luckysheetfile[sheetIndex];

                    file.row = d.length;
                    file.data = d;

                    jfrefreshgridall(
                        d[0].length,
                        d.length,
                        d,
                        null,
                        Store.luckysheet_select_save,
                        'datachangeAll',
                        undefined,
                        undefined
                    );
                } else {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    Store.luckysheetfile[sheetIndex].data = d;
                }

            }, () => {
                return formula.error.v;
            });

            return 'loading...';
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'ASK_AI': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const luckysheetCurrentIndex = window.luckysheetCurrentIndex;
            const currentSheetIndex = Store.currentSheetIndex;

            // 工作表不匹配直接返回
            if (luckysheetCurrentIndex !== currentSheetIndex) {
                return;
            }

            const startRow = window.luckysheetCurrentRow;
            const startColumn = window.luckysheetCurrentColumn;
            const cellFp = window.luckysheetCurrentFunction;

            // 获取参数
            const targetText = func_methods.getFirstValue(arguments[0]);
            let rangeData;

            // 获取数据源
            if (arguments[1]) {
                rangeData = formula.getRangeArrayTwo(arguments[1].data);
            } else {
                rangeData = excelToArray(companyTargetData);
            }

            const companyTarget = excelToArray(companyTargetData);
            let resultTable = askAIData(rangeData, companyTarget);

            // 未传入区域时，使用月份默认数据
            if (!arguments[1]) {
                if (targetText.includes('10月')) {
                    resultTable = excelToLuckyArray(companyTargetData10);
                } else if (targetText.includes('11月')) {
                    resultTable = excelToLuckyArray(companyTargetData11);
                } else if (targetText.includes('12月')) {
                    resultTable = excelToLuckyArray(companyTargetData12);
                } else {
                    resultTable = excelToLuckyArray(companyTargetData11);
                }
            }

            setTimeout(() => {
                let d = editor.deepCopyFlowData(Store.flowdata);

                const targetRows = resultTable.length;
                const targetCols = resultTable[0]?.length || 0;
                const rowHeight = startRow + targetRows;
                const colWidth = startColumn + targetCols;

                // 自动扩展表格
                if (rowHeight >= d.length && colWidth >= d[0].length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, colWidth - d[0].length + 1);
                } else if (rowHeight >= d.length) {
                    d = datagridgrowth(d, rowHeight - d.length + 1, 0);
                } else if (colWidth >= d[0].length) {
                    d = datagridgrowth(d, 0, colWidth - d[0].length + 1);
                }

                // 填充数据
                resultTable.forEach((row, r) => {
                    row.forEach((cell, c) => {
                        setcellvalue(startRow + r, startColumn + c, d, cell);
                    });
                });

                // 恢复公式
                d[startRow][startColumn].f = cellFp;
                delete d[startRow][startColumn].m;

                // 刷新表格
                if (currentSheetIndex === Store.currentSheetIndex) {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    const file = Store.luckysheetfile[sheetIndex];

                    file.row = d.length;
                    file.data = d;

                    jfrefreshgridall(
                        d[0].length,
                        d.length,
                        d,
                        null,
                        Store.luckysheet_select_save,
                        'datachangeAll',
                        undefined,
                        undefined
                    );
                } else {
                    const sheetIndex = getSheetIndex(Store.currentSheetIndex);
                    Store.luckysheetfile[sheetIndex].data = d;
                }

            }, 300);

            return 'loading...';
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'COUNTIFS': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const args = arguments;
            luckysheet_getValue(args);

            // 初始化条件匹配结果
            const firstRange = formula.getRangeArray(args[0])[0];
            const conditionResults = new Array(firstRange.length).fill(true);

            // 遍历所有条件组
            for (let i = 0; i < args.length; i += 2) {
                const range = formula.getRangeArray(args[i])[0];
                const criteria = args[i + 1];

                for (let j = 0; j < range.length; j++) {
                    const value = range[j];
                    conditionResults[j] = conditionResults[j] && !!value && formula.acompareb(value, criteria);
                }
            }

            // 统计满足条件的数量
            let result = 0;
            for (let i = 0; i < conditionResults.length; i++) {
                if (conditionResults[i]) {
                    result++;
                }
            }

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PRODUCT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历所有参数，展开区域/数组
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    dataArr.push(data);
                }
            }

            // 乘积计算
            let result = 1;
            for (let i = 0; i < dataArr.length; i++) {
                let number = dataArr[i];

                if (!isRealNum(number)) {
                    return formula.error.v;
                }

                number = parseFloat(number);
                result *= number;
            }

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'HARMEAN': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    dataArr.push(data);
                }
            }

            let den = 0;
            let len = 0;

            // 计算调和平均数
            for (let i = 0; i < dataArr.length; i++) {
                let number = dataArr[i];

                if (!isRealNum(number)) {
                    return formula.error.v;
                }

                number = parseFloat(number);

                // 数值必须大于 0
                if (number <= 0) {
                    return formula.error.nm;
                }

                den += 1 / number;
                len++;
            }

            return len / den;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'HYPGEOMDIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 样本中成功的次数
            let sample_s = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(sample_s)) return sample_s;
            if (!isRealNum(sample_s)) return formula.error.v;
            sample_s = parseInt(sample_s);

            // 样本量
            let number_sample = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(number_sample)) return number_sample;
            if (!isRealNum(number_sample)) return formula.error.v;
            number_sample = parseInt(number_sample);

            // 总体中成功的次数
            let population_s = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(population_s)) return population_s;
            if (!isRealNum(population_s)) return formula.error.v;
            population_s = parseInt(population_s);

            // 总体大小
            let number_pop = func_methods.getFirstValue(arguments[3]);
            if (valueIsError(number_pop)) return number_pop;
            if (!isRealNum(number_pop)) return formula.error.v;
            number_pop = parseInt(number_pop);

            // 决定函数形式的逻辑值
            const cumulative = func_methods.getCellBoolen(arguments[4]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            const minVal = Math.min(number_sample, population_s);
            const maxVal = Math.max(0, number_sample - number_pop + population_s);
            if (sample_s < 0 || sample_s > minVal || sample_s < maxVal) {
                return formula.error.nm;
            }
            if (number_sample <= 0 || number_sample > number_pop) return formula.error.nm;
            if (population_s <= 0 || population_s > number_pop) return formula.error.nm;
            if (number_pop <= 0) return formula.error.nm;

            // 概率密度函数
            function pdf(x, n, M, N) {
                const a = func_methods.factorial(M) / (func_methods.factorial(x) * func_methods.factorial(M - x));
                const b = func_methods.factorial(N - M) / (func_methods.factorial(n - x) * func_methods.factorial(N - M - n + x));
                const c = func_methods.factorial(N) / (func_methods.factorial(n) * func_methods.factorial(N - n));
                return a * b / c;
            }

            // 累积分布函数
            function cdf(x, n, M, N) {
                let sum = 0;
                for (let i = 0; i <= x; i++) {
                    sum += pdf(i, n, M, N);
                }
                return sum;
            }

            return cumulative ? cdf(sample_s, number_sample, population_s, number_pop) : pdf(sample_s, number_sample, population_s, number_pop);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'INTERCEPT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // x轴上用于预测的值（截距固定为0）
            const x = 0;

            // 处理因变量 known_y's
            const data_known_y = arguments[0];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 处理自变量 known_x's
            const data_known_x = arguments[1];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 数据长度必须一致
            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            // 筛选有效数值
            const data_y = [], data_x = [];
            for (let i = 0; i < known_y.length; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(parseFloat(num_y));
                    data_x.push(parseFloat(num_x));
                }
            }

            // 方差为0时返回错误
            if (func_methods.variance_s(data_x) === 0) {
                return formula.error.d;
            }

            // 计算截距
            const xmean = jStat.mean(data_x);
            const ymean = jStat.mean(data_y);
            const n = data_x.length;

            let num = 0;
            let den = 0;

            for (let i = 0; i < n; i++) {
                num += (data_x[i] - xmean) * (data_y[i] - ymean);
                den += Math.pow(data_x[i] - xmean, 2);
            }

            const b = num / den;
            const a = ymean - b * xmean;

            return a + b * x;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'KURT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
                } else {
                    dataArr.push(data);
                }
            }

            // 剔除不是数值类型的值
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                let number = dataArr[j];

                if (!isRealNum(number)) {
                    return formula.error.v;
                }

                number = parseFloat(number);
                dataArr_n.push(number);
            }

            // 数据校验：至少4个数值，且标准差不能为0
            if (dataArr_n.length < 4 || func_methods.standardDeviation_s(dataArr_n) === 0) {
                return formula.error.d;
            }

            // 计算峰度 KURT
            const mean = jStat.mean(dataArr_n);
            const stdev = jStat.stdev(dataArr_n, true);
            const n = dataArr_n.length;

            let sigma = 0;
            for (let i = 0; i < n; i++) {
                sigma += Math.pow(dataArr_n[i] - mean, 4);
            }
            sigma = sigma / Math.pow(stdev, 4);

            const part1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3));
            const part2 = (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));

            return part1 * sigma - part2;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'LARGE': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 处理数据区域/数组
            let dataArr = [];
            const dataArg = arguments[0];

            if (getObjType(dataArg) === 'array') {
                if (getObjType(dataArg[0]) === 'array' && !func_methods.isDyadicArr(dataArg)) {
                    return formula.error.v;
                }
                dataArr = dataArr.concat(func_methods.getDataArr(dataArg, true));
            } else if (getObjType(dataArg) === 'object' && dataArg.startCell != null) {
                dataArr = dataArr.concat(func_methods.getCellDataArr(dataArg, 'text', true));
            } else {
                dataArr.push(dataArg);
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                const number = dataArr[j];
                if (!isRealNum(number)) {
                    return formula.error.v;
                }
                dataArr_n.push(parseFloat(number));
            }

            // 处理第 K 大参数
            let n;
            const kArg = arguments[1];

            if (getObjType(kArg) === 'array') {
                if (getObjType(kArg[0]) === 'array' && !func_methods.isDyadicArr(kArg)) {
                    return formula.error.v;
                }
                n = func_methods.getDataArr(kArg);
            } else if (getObjType(kArg) === 'object' && kArg.startCell != null) {
                if (kArg.rowl > 1 || kArg.coll > 1) {
                    return formula.error.v;
                }
                const cell = kArg.data;
                n = (cell == null || isRealNull(cell.v)) ? 0 : cell.v;
            } else {
                n = kArg;
            }

            // 排序（降序）
            const sortedArr = [...dataArr_n].sort((a, b) => b - a);

            // 数组模式返回多个结果
            if (getObjType(n) === 'array') {
                if (sortedArr.length === 0) return formula.error.nm;
                const result = [];

                for (let i = 0; i < n.length; i++) {
                    const k = n[i];
                    if (!isRealNum(k)) {
                        result.push(formula.error.v);
                        continue;
                    }
                    const kVal = Math.ceil(parseFloat(k));
                    if (kVal <= 0 || kVal > sortedArr.length) {
                        result.push(formula.error.nm);
                        continue;
                    }
                    result.push(sortedArr[kVal - 1]);
                }
                return result;
            }

            // 单个值模式
            if (!isRealNum(n)) return formula.error.v;
            const kVal = Math.ceil(parseFloat(n));

            if (sortedArr.length === 0 || kVal <= 0 || kVal > sortedArr.length) {
                return formula.error.nm;
            }

            return sortedArr[kVal - 1];
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'STDEVA': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, false));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', false));
                } else {
                    dataArr.push(data);
                }
            }

            // 非数值类型转换规则：true→1，其他→0
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                let number = dataArr[j];

                if (!isRealNum(number)) {
                    if (String(number).toLowerCase() === 'true') {
                        number = 1;
                    } else {
                        number = 0;
                    }
                } else {
                    number = parseFloat(number);
                }

                dataArr_n.push(number);
            }

            // 数据校验
            if (dataArr_n.length === 0) {
                return 0;
            }
            if (dataArr_n.length === 1) {
                return formula.error.d;
            }

            return func_methods.standardDeviation_s(dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'STDEVP': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
                } else {
                    dataArr.push(data);
                }
            }

            // 剔除不是数值类型的值
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                let number = dataArr[j];

                if (!isRealNum(number)) {
                    return formula.error.v;
                }

                number = parseFloat(number);
                dataArr_n.push(number);
            }

            // 数据校验
            if (dataArr_n.length === 0) {
                return 0;
            }
            if (dataArr_n.length === 1) {
                return formula.error.d;
            }

            return func_methods.standardDeviation(dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'GEOMEAN': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', true));
                } else {
                    if (getObjType(data) === 'boolean') {
                        dataArr.push(data ? 1 : 0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            // 筛选有效数值，必须大于 0
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                let number = dataArr[j];

                if (!isRealNum(number)) {
                    continue;
                }

                number = parseFloat(number);

                if (number <= 0) {
                    return formula.error.nm;
                }

                dataArr_n.push(number);
            }

            // 无有效数据返回错误
            if (dataArr_n.length === 0) {
                return formula.error.nm;
            }

            return jStat.geomean(dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'RANK_EQ': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要确定其排名的值
            let number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number)) {
                return number;
            }
            if (!isRealNum(number)) {
                return formula.error.v;
            }
            number = parseFloat(number);

            // 包含相关数据集的数组或范围
            const data_ref = arguments[1];
            let ref = [];

            if (getObjType(data_ref) === 'array') {
                if (getObjType(data_ref[0]) === 'array' && !func_methods.isDyadicArr(data_ref)) {
                    return formula.error.v;
                }
                ref = ref.concat(func_methods.getDataArr(data_ref, true));
            } else if (getObjType(data_ref) === 'object' && data_ref.startCell != null) {
                ref = ref.concat(func_methods.getCellDataArr(data_ref, 'number', true));
            } else {
                ref.push(data_ref);
            }

            // 筛选数值
            const ref_n = [];
            for (let j = 0; j < ref.length; j++) {
                const num = ref[j];
                if (!isRealNum(num)) {
                    return formula.error.v;
                }
                ref_n.push(parseFloat(num));
            }

            // 排序顺序：0/省略=降序，非0=升序
            let order = false;
            if (arguments.length === 3) {
                order = func_methods.getCellBoolen(arguments[2]);
                if (valueIsError(order)) {
                    return order;
                }
            }

            // 排序
            const sortedRef = [...ref_n].sort(order ? (a, b) => a - b : (a, b) => b - a);

            // 计算排名（相同数值返回相同排名，符合 RANK.EQ 标准）
            const rank = sortedRef.findIndex(item => item === number) + 1;

            if (rank === 0) {
                return formula.error.na;
            }

            return rank;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'RANK_AVG': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要确定其排名的值
            let number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number)) {
                return number;
            }
            if (!isRealNum(number)) {
                return formula.error.v;
            }
            number = parseFloat(number);

            // 包含相关数据集的数组或范围
            const data_ref = arguments[1];
            let ref = [];

            if (getObjType(data_ref) === 'array') {
                if (getObjType(data_ref[0]) === 'array' && !func_methods.isDyadicArr(data_ref)) {
                    return formula.error.v;
                }
                ref = ref.concat(func_methods.getDataArr(data_ref, true));
            } else if (getObjType(data_ref) === 'object' && data_ref.startCell != null) {
                ref = ref.concat(func_methods.getCellDataArr(data_ref, 'number', true));
            } else {
                ref.push(data_ref);
            }

            // 筛选数值
            const ref_n = [];
            for (let j = 0; j < ref.length; j++) {
                const num = ref[j];
                if (!isRealNum(num)) {
                    return formula.error.v;
                }
                ref_n.push(parseFloat(num));
            }

            // 排序顺序：0=降序，非0=升序
            let order = false;
            if (arguments.length === 3) {
                order = func_methods.getCellBoolen(arguments[2]);
                if (valueIsError(order)) {
                    return order;
                }
            }

            // 排序（不修改原数组）
            const sortedRef = [...ref_n].sort(order ? (a, b) => a - b : (a, b) => b - a);

            // 查找第一个匹配位置
            const firstIndex = sortedRef.findIndex(item => item === number);
            if (firstIndex === -1) {
                return formula.error.na;
            }

            // 统计重复值数量
            let count = 0;
            for (let i = 0; i < sortedRef.length; i++) {
                if (sortedRef[i] === number) {
                    count++;
                }
            }

            // 计算平均排名
            return count > 1 ? (2 * firstIndex + count + 1) / 2 : firstIndex + 1;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PERCENTRANK_EXC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 包含相关数据集的数组或范围
            const data_ref = arguments[0];
            let ref = [];

            if (getObjType(data_ref) === 'array') {
                if (getObjType(data_ref[0]) === 'array' && !func_methods.isDyadicArr(data_ref)) {
                    return formula.error.v;
                }
                ref = ref.concat(func_methods.getDataArr(data_ref, true));
            } else if (getObjType(data_ref) === 'object' && data_ref.startCell != null) {
                ref = ref.concat(func_methods.getCellDataArr(data_ref, 'number', true));
            } else {
                ref.push(data_ref);
            }

            // 筛选有效数值
            const ref_n = [];
            for (let j = 0; j < ref.length; j++) {
                let number = ref[j];
                if (!isRealNum(number)) {
                    return formula.error.v;
                }
                ref_n.push(parseFloat(number));
            }

            // 要确定其百分比排位的值
            let x = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(x)) {
                return x;
            }
            if (!isRealNum(x)) {
                return formula.error.v;
            }
            x = parseFloat(x);

            // 有效位数
            let significance = 3;
            if (arguments.length === 3) {
                significance = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(significance)) {
                    return significance;
                }
                if (!isRealNum(significance)) {
                    return formula.error.v;
                }
                significance = parseInt(significance);
            }

            // 基础校验
            if (ref_n.length === 0) {
                return formula.error.nm;
            }
            if (significance < 1) {
                return formula.error.nm;
            }

            // 特殊情况：只有一个数据
            if (ref_n.length === 1 && ref_n[0] === x) {
                return 1;
            }

            // 升序排序
            const sortedRef = [...ref_n].sort((a, b) => a - b);
            const uniques = window.luckysheet_function.UNIQUE.f(sortedRef)[0];

            const n = sortedRef.length;
            const m = uniques.length;
            const power = Math.pow(10, significance);

            let result = 0;
            let match = false;
            let i = 0;

            // 计算排位
            while (!match && i < m) {
                if (x === uniques[i]) {
                    result = (sortedRef.indexOf(uniques[i]) + 1) / (n + 1);
                    match = true;
                } else if (x >= uniques[i] && (x < uniques[i + 1] || i === m - 1)) {
                    result = (sortedRef.lastIndexOf(uniques[i]) + 1 + (x - uniques[i]) / (uniques[i + 1] - uniques[i])) / (n + 1);
                    match = true;
                }
                i++;
            }

            // 结果处理
            if (isNaN(result)) {
                return formula.error.na;
            } else {
                return Math.floor(result * power) / power;
            }
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PERCENTRANK_INC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 包含相关数据集的数组或范围
            const data_ref = arguments[0];
            let ref = [];

            if (getObjType(data_ref) === 'array') {
                if (getObjType(data_ref[0]) === 'array' && !func_methods.isDyadicArr(data_ref)) {
                    return formula.error.v;
                }
                ref = ref.concat(func_methods.getDataArr(data_ref, true));
            } else if (getObjType(data_ref) === 'object' && data_ref.startCell != null) {
                ref = ref.concat(func_methods.getCellDataArr(data_ref, 'number', true));
            } else {
                ref.push(data_ref);
            }

            // 筛选有效数值
            const ref_n = [];
            for (let j = 0; j < ref.length; j++) {
                let number = ref[j];
                if (!isRealNum(number)) {
                    return formula.error.v;
                }
                ref_n.push(parseFloat(number));
            }

            // 要确定其百分比排位的值
            let x = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(x)) {
                return x;
            }
            if (!isRealNum(x)) {
                return formula.error.v;
            }
            x = parseFloat(x);

            // 有效位数
            let significance = 3;
            if (arguments.length === 3) {
                significance = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(significance)) {
                    return significance;
                }
                if (!isRealNum(significance)) {
                    return formula.error.v;
                }
                significance = parseInt(significance);
            }

            // 基础校验
            if (ref_n.length === 0) {
                return formula.error.nm;
            }
            if (significance < 1) {
                return formula.error.nm;
            }

            // 特殊情况：只有一个数据
            if (ref_n.length === 1 && ref_n[0] === x) {
                return 1;
            }

            // 升序排序
            const sortedRef = [...ref_n].sort((a, b) => a - b);
            const uniques = window.luckysheet_function.UNIQUE.f(sortedRef)[0];

            const n = sortedRef.length;
            const m = uniques.length;
            const power = Math.pow(10, significance);

            let result = 0;
            let match = false;
            let i = 0;

            // 计算排位（包含 0 和 1，标准 Excel PERCENTRANK.INC 逻辑）
            while (!match && i < m) {
                if (x === uniques[i]) {
                    result = sortedRef.indexOf(uniques[i]) / (n - 1);
                    match = true;
                } else if (x >= uniques[i] && (x < uniques[i + 1] || i === m - 1)) {
                    result = (sortedRef.lastIndexOf(uniques[i]) + (x - uniques[i]) / (uniques[i + 1] - uniques[i])) / (n - 1);
                    match = true;
                }
                i++;
            }

            // 结果处理
            if (isNaN(result)) {
                return formula.error.na;
            } else {
                return Math.floor(result * power) / power;
            }
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'FORECAST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // x轴上用于预测的值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) {
                return x;
            }
            if (!isRealNum(x)) {
                return formula.error.v;
            }
            x = parseFloat(x);

            // 代表因变量数据数组或矩阵的范围
            const data_known_y = arguments[1];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 代表自变量数据数组或矩阵的范围
            const data_known_x = arguments[2];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 数据长度必须一致
            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            // known_y 和 known_x 只取数值
            const data_y = [], data_x = [];
            for (let i = 0; i < known_y.length; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(parseFloat(num_y));
                    data_x.push(parseFloat(num_x));
                }
            }

            // 方差为0时返回错误
            if (func_methods.variance_s(data_x) === 0) {
                return formula.error.d;
            }

            // 线性回归计算
            const xmean = jStat.mean(data_x);
            const ymean = jStat.mean(data_y);
            const n = data_x.length;

            let num = 0;
            let den = 0;

            for (let i = 0; i < n; i++) {
                num += (data_x[i] - xmean) * (data_y[i] - ymean);
                den += Math.pow(data_x[i] - xmean, 2);
            }

            const b = num / den;
            const a = ymean - b * xmean;

            return a + b * x;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'FISHERINV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let y = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(y)) {
                return y;
            }

            if (!isRealNum(y)) {
                return formula.error.v;
            }

            y = parseFloat(y);
            const e2y = Math.exp(2 * y);

            return (e2y - 1) / (e2y + 1);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'FISHER': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) {
                return x;
            }

            if (!isRealNum(x)) {
                return formula.error.v;
            }

            x = parseFloat(x);

            if (x <= -1 || x >= 1) {
                return formula.error.nm;
            }

            return Math.log((1 + x) / (1 - x)) / 2;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MODE_SNGL': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 统计频次，寻找众数
            const count = {};
            let maxItems = [];
            let max = 0;
            let currentItem;

            for (let i = 0; i < dataArr_n.length; i++) {
                currentItem = dataArr_n[i];
                count[currentItem] = (count[currentItem] || 0) + 1;

                if (count[currentItem] > max) {
                    max = count[currentItem];
                    maxItems = [];
                }

                if (count[currentItem] === max) {
                    maxItems.push(currentItem);
                }
            }

            // 无众数（频次都为1）
            if (max <= 1) {
                return formula.error.na;
            }

            // 找到最早出现的众数
            let resultIndex = dataArr_n.indexOf(maxItems[0]);
            for (let j = 0; j < maxItems.length; j++) {
                const index = dataArr_n.indexOf(maxItems[j]);
                if (index < resultIndex) {
                    resultIndex = index;
                }
            }

            return dataArr_n[resultIndex];
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'WEIBULL_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // WEIBULL 分布函数的输入值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) return formula.error.v;
            x = parseFloat(x);

            // Weibull 分布函数的形状参数
            let alpha = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(alpha)) return alpha;
            if (!isRealNum(alpha)) return formula.error.v;
            alpha = parseFloat(alpha);

            // Weibull 分布函数的尺度参数
            let beta = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(beta)) return beta;
            if (!isRealNum(beta)) return formula.error.v;
            beta = parseFloat(beta);

            // 决定函数形式的逻辑值
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (x < 0 || alpha <= 0 || beta <= 0) {
                return formula.error.nm;
            }

            // 计算分布结果
            return cumulative
                ? 1 - Math.exp(-Math.pow(x / beta, alpha))
                : Math.pow(x, alpha - 1) * Math.exp(-Math.pow(x / beta, alpha)) * alpha / Math.pow(beta, alpha);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'AVEDEV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 数据校验
            if (dataArr_n.length === 0) {
                return formula.error.nm;
            }

            // 计算平均绝对偏差
            const mean = jStat.mean(dataArr_n);
            const sumAbsDev = jStat.sum(jStat(dataArr_n).subtract(mean).abs()[0]);

            return sumAbsDev / dataArr_n.length;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'AVERAGEA': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, false));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    // 布尔值处理
                    if (getObjType(data) === 'boolean') {
                        dataArr.push(data ? 1 : 0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            let sum = 0;
            let count = 0;

            // 计算总和与计数
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];

                if (isRealNum(number)) {
                    sum += parseFloat(number);
                } else {
                    // 非数值：true=1，其他=0
                    sum += String(number).toLowerCase() === 'true' ? 1 : 0;
                }
                count++;
            }

            if (count === 0) {
                return formula.error.d;
            }

            return sum / count;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'BINOM_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 试验的成功次数
            let number_s = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number_s)) return number_s;
            if (!isRealNum(number_s)) return formula.error.v;
            number_s = parseInt(number_s);

            // 独立试验的次数
            let trials = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(trials)) return trials;
            if (!isRealNum(trials)) return formula.error.v;
            trials = parseInt(trials);

            // 单次试验成功概率
            let probability_s = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(probability_s)) return probability_s;
            if (!isRealNum(probability_s)) return formula.error.v;
            probability_s = parseFloat(probability_s);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (number_s < 0 || number_s > trials) return formula.error.nm;
            if (trials < 0) return formula.error.nm;
            if (probability_s < 0 || probability_s > 1) return formula.error.nm;

            // 计算二项分布
            return cumulative
                ? jStat.binomial.cdf(number_s, trials, probability_s)
                : jStat.binomial.pdf(number_s, trials, probability_s);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'BINOM_INV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 贝努利试验次数
            let trials = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(trials)) return trials;
            if (!isRealNum(trials)) return formula.error.v;
            trials = parseInt(trials);

            // 单次试验成功概率
            let probability_s = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(probability_s)) return probability_s;
            if (!isRealNum(probability_s)) return formula.error.v;
            probability_s = parseFloat(probability_s);

            // 期望的临界概率
            let alpha = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(alpha)) return alpha;
            if (!isRealNum(alpha)) return formula.error.v;
            alpha = parseFloat(alpha);

            // 参数合法性校验
            if (trials < 0) return formula.error.nm;
            if (probability_s < 0 || probability_s > 1) return formula.error.nm;
            if (alpha < 0 || alpha > 1) return formula.error.nm;

            // 计算最小成功次数
            let x = 0;
            while (x <= trials) {
                if (jStat.binomial.cdf(x, trials, probability_s) >= alpha) {
                    return x;
                }
                x++;
            }

            return trials;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'CONFIDENCE_NORM': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 置信水平
            let alpha = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(alpha)) return alpha;
            if (!isRealNum(alpha)) return formula.error.v;
            alpha = parseFloat(alpha);

            // 数据区域的总体标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(standard_dev)) return standard_dev;
            if (!isRealNum(standard_dev)) return formula.error.v;
            standard_dev = parseFloat(standard_dev);

            // 样本总量的大小
            let size = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(size)) return size;
            if (!isRealNum(size)) return formula.error.v;
            size = parseInt(size);

            // 参数合法性校验
            if (alpha <= 0 || alpha >= 1) return formula.error.nm;
            if (standard_dev <= 0) return formula.error.nm;
            if (size < 1) return formula.error.nm;

            // 计算置信区间
            return jStat.normalci(1, alpha, standard_dev, size)[1] - 1;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'CORREL': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 处理因变量数据 known_y's
            const data_known_y = arguments[0];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 处理自变量数据 known_x's
            const data_known_x = arguments[1];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 数据长度必须一致
            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            // 筛选有效数值
            const data_y = [], data_x = [];
            for (let i = 0; i < known_y.length; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(parseFloat(num_y));
                    data_x.push(parseFloat(num_x));
                }
            }

            // 数据有效性校验
            if (data_y.length === 0 || data_x.length === 0 ||
                func_methods.standardDeviation(data_y) === 0 ||
                func_methods.standardDeviation(data_x) === 0) {
                return formula.error.d;
            }

            // 计算相关系数
            return jStat.corrcoeff(data_y, data_x);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'COVARIANCE_P': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 自变量数据数组
            const data_known_x = arguments[0];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 因变量数据数组
            const data_known_y = arguments[1];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 数据长度必须一致
            if (known_x.length !== known_y.length) {
                return formula.error.na;
            }

            // 筛选有效数值
            const data_x = [], data_y = [];
            for (let i = 0; i < known_x.length; i++) {
                const num_x = known_x[i];
                const num_y = known_y[i];

                if (isRealNum(num_x) && isRealNum(num_y)) {
                    data_x.push(parseFloat(num_x));
                    data_y.push(parseFloat(num_y));
                }
            }

            // 数据有效性校验
            if (data_x.length === 0 || data_y.length === 0) {
                return formula.error.d;
            }

            // 计算总体协方差
            const mean_x = jStat.mean(data_x);
            const mean_y = jStat.mean(data_y);
            let sum = 0;

            for (let i = 0; i < data_x.length; i++) {
                sum += (data_x[i] - mean_x) * (data_y[i] - mean_y);
            }

            return sum / data_x.length;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'COVARIANCE_S': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 自变量数据数组
            const data_known_x = arguments[0];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 因变量数据数组
            const data_known_y = arguments[1];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 数据长度必须一致
            if (known_x.length !== known_y.length) {
                return formula.error.na;
            }

            // 筛选有效数值
            const data_x = [], data_y = [];
            for (let i = 0; i < known_x.length; i++) {
                const num_x = known_x[i];
                const num_y = known_y[i];

                if (isRealNum(num_x) && isRealNum(num_y)) {
                    data_x.push(parseFloat(num_x));
                    data_y.push(parseFloat(num_y));
                }
            }

            // 数据有效性校验
            if (data_x.length === 0 || data_y.length === 0) {
                return formula.error.d;
            }

            // 计算样本协方差
            return jStat.covariance(data_x, data_y);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DEVSQ': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历参数，展开数组/区域数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        // 布尔值处理
                        if (getObjType(data) === 'boolean') {
                            dataArr.push(data ? 1 : 0);
                        } else {
                            return formula.error.v;
                        }
                    } else {
                        dataArr.push(data);
                    }
                }
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 计算平均偏差平方和
            const mean = jStat.mean(dataArr_n);
            let result = 0;
            for (let i = 0; i < dataArr_n.length; i++) {
                result += Math.pow(dataArr_n[i] - mean, 2);
            }

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'EXPON_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 指数分布函数的输入值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) return formula.error.v;
            x = parseFloat(x);

            // 指数分布的参数 lambda
            let lambda = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(lambda)) return lambda;
            if (!isRealNum(lambda)) return formula.error.v;
            lambda = parseFloat(lambda);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[2]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (x < 0 || lambda <= 0) return formula.error.nm;

            // 计算指数分布
            return cumulative
                ? jStat.exponential.cdf(x, lambda)
                : jStat.exponential.pdf(x, lambda);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'AVERAGEIF': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let sum = 0;
            let count = 0;

            // 条件判断区域
            let rangeData = arguments[0].data;
            const rangeRow = arguments[0].rowl;
            const rangeCol = arguments[0].coll;
            const criteria = luckysheet_parseData(arguments[1]);
            let sumRangeData = [];

            // 处理第三个参数：求和区域
            if (arguments[2]) {
                const sumRangeStart = arguments[2].startCell;
                const sumRangeRow = arguments[2].rowl;
                const sumRangeCol = arguments[2].coll;
                const sumRangeSheet = arguments[2].sheetName;

                if (rangeRow === sumRangeRow && rangeCol === sumRangeCol) {
                    sumRangeData = arguments[2].data;
                } else {
                    // 计算实际求和区域
                    const row = [];
                    const col = [];
                    let sumRangeEnd = '';
                    let realSumRange = '';

                    row[0] = parseInt(sumRangeStart.replace(/[^0-9]/g, '')) - 1;
                    col[0] = ABCatNum(sumRangeStart.replace(/[^A-Za-z]/g, ''));

                    // 按条件区域尺寸扩展
                    row[1] = row[0] + rangeRow - 1;
                    col[1] = col[0] + rangeCol - 1;

                    // 转换为单元格地址
                    const real_ABC = chatatABC(col[1]);
                    const real_Num = row[1] + 1;
                    sumRangeEnd = real_ABC + real_Num;

                    realSumRange = `${sumRangeSheet}!${sumRangeStart}:${sumRangeEnd}`;
                    sumRangeData = luckysheet_getcelldata(realSumRange).data;
                }

                sumRangeData = formula.getRangeArray(sumRangeData)[0];
            }

            rangeData = formula.getRangeArray(rangeData)[0];

            // 遍历匹配并计算
            for (let i = 0; i < rangeData.length; i++) {
                const v = rangeData[i];
                if (!!v && formula.acompareb(v, criteria)) {
                    const vnow = sumRangeData[i] || v;

                    if (!isRealNum(vnow)) {
                        continue;
                    }

                    sum += parseFloat(vnow);
                    count++;
                }
            }

            // 无有效数据返回错误
            if (count === 0) {
                return formula.error.d;
            }

            return numFormat(sum / count);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'AVERAGEIFS': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let sum = 0;
            let count = 0;
            const args = arguments;

            luckysheet_getValue(args);

            // 平均值计算区域
            const rangeData = formula.getRangeArray(args[0])[0];
            // 初始化所有行匹配结果为 true
            const results = new Array(rangeData.length).fill(true);

            // 遍历多组条件区域和条件
            for (let i = 1; i < args.length; i += 2) {
                const range = formula.getRangeArray(args[i])[0];
                const criteria = args[i + 1];

                // 逐行判断是否满足当前条件
                for (let j = 0; j < range.length; j++) {
                    const v = range[j];
                    results[j] = results[j] && !!v && formula.acompareb(v, criteria);
                }
            }

            // 统计满足所有条件的有效数值
            for (let i = 0; i < rangeData.length; i++) {
                if (results[i] && isRealNum(rangeData[i])) {
                    sum += parseFloat(rangeData[i]);
                    count++;
                }
            }

            // 无有效数据时返回错误（符合 Excel 标准）
            if (count === 0) {
                return formula.error.d;
            }

            return numFormat(sum / count);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PERMUT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 表示对象总个数的整数
            let number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number)) return number;
            if (!isRealNum(number)) return formula.error.v;
            number = parseInt(number);

            // 表示每个排列中对象个数的整数
            let number_chosen = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(number_chosen)) return number_chosen;
            if (!isRealNum(number_chosen)) return formula.error.v;
            number_chosen = parseInt(number_chosen);

            // 参数合法性校验
            if (number <= 0 || number_chosen < 0 || number < number_chosen) {
                return formula.error.nm;
            }

            // 计算排列数：P(n,k) = n!/(n-k)!
            return func_methods.factorial(number) / func_methods.factorial(number - number_chosen);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'TRIMMEAN': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 数据集处理
            const data_dataArr = arguments[0];
            let dataArr = [];

            if (getObjType(data_dataArr) === 'array') {
                if (getObjType(data_dataArr[0]) === 'array' && !func_methods.isDyadicArr(data_dataArr)) {
                    return formula.error.v;
                }
                dataArr = dataArr.concat(func_methods.getDataArr(data_dataArr, false));
            } else if (getObjType(data_dataArr) === 'object' && data_dataArr.startCell != null) {
                dataArr = dataArr.concat(func_methods.getCellDataArr(data_dataArr, 'number', false));
            } else {
                dataArr.push(data_dataArr);
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 排除比例
            let percent = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(percent)) return percent;
            if (!isRealNum(percent)) return formula.error.v;
            percent = parseFloat(percent);

            // 参数合法性校验
            if (dataArr_n.length === 0 || percent < 0 || percent > 1) {
                return formula.error.nm;
            }

            // 升序排序
            dataArr_n.sort((a, b) => a - b);

            // 计算需要修剪的数量（Excel 规则：向下取偶数）
            const trimCount = Math.floor(dataArr_n.length * percent / 2) * 2;
            const exclude = trimCount / 2;

            // 修剪首尾并计算平均值
            const trimmedArray = dataArr_n.slice(exclude, dataArr_n.length - exclude);
            const result = jStat.mean(trimmedArray);

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PERCENTILE_EXC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 定义相对位置的数组或数据区域
            const data_dataArr = arguments[0];
            let dataArr = [];

            if (getObjType(data_dataArr) === 'array') {
                if (getObjType(data_dataArr[0]) === 'array' && !func_methods.isDyadicArr(data_dataArr)) {
                    return formula.error.v;
                }
                dataArr = dataArr.concat(func_methods.getDataArr(data_dataArr, false));
            } else if (getObjType(data_dataArr) === 'object' && data_dataArr.startCell != null) {
                dataArr = dataArr.concat(func_methods.getCellDataArr(data_dataArr, 'number', false));
            } else {
                dataArr.push(data_dataArr);
            }

            // 筛选有效数值
            let dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 百分点值（0~1之间，不含0和1）
            let k = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(k)) return k;
            if (!isRealNum(k)) return formula.error.v;
            k = parseFloat(k);

            // 参数合法性校验
            const n = dataArr_n.length;
            if (n === 0) return formula.error.nm;
            if (k <= 0 || k >= 1) return formula.error.nm;
            if (k < 1 / (n + 1) || k > 1 - 1 / (n + 1)) return formula.error.nm;

            // 升序排序
            dataArr_n.sort((a, b) => a - b);

            // 计算百分比结果（Excel 标准算法）
            const l = k * (n + 1) - 1;
            const fl = Math.floor(l);

            return l === fl ? dataArr_n[l] : dataArr_n[fl] + (l - fl) * (dataArr_n[fl + 1] - dataArr_n[fl]);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PERCENTILE_INC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 定义相对位置的数组或数据区域
            const data_dataArr = arguments[0];
            let dataArr = [];

            if (getObjType(data_dataArr) === 'array') {
                if (getObjType(data_dataArr[0]) === 'array' && !func_methods.isDyadicArr(data_dataArr)) {
                    return formula.error.v;
                }
                dataArr = dataArr.concat(func_methods.getDataArr(data_dataArr, false));
            } else if (getObjType(data_dataArr) === 'object' && data_dataArr.startCell != null) {
                dataArr = dataArr.concat(func_methods.getCellDataArr(data_dataArr, 'number', false));
            } else {
                dataArr.push(data_dataArr);
            }

            // 筛选有效数值
            let dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 百分点值（0~1之间，包含0和1）
            let k = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(k)) return k;
            if (!isRealNum(k)) return formula.error.v;
            k = parseFloat(k);

            // 参数合法性校验
            if (dataArr_n.length === 0 || k < 0 || k > 1) {
                return formula.error.nm;
            }

            // 升序排序
            dataArr_n.sort((a, b) => a - b);

            // 计算百分比结果（Excel 包含型标准算法）
            const n = dataArr_n.length;
            const l = k * (n - 1);
            const fl = Math.floor(l);

            return l === fl ? dataArr_n[l] : dataArr_n[fl] + (l - fl) * (dataArr_n[fl + 1] - dataArr_n[fl]);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PEARSON': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 自变量数据数组
            const data_known_x = arguments[0];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'text', false));
            } else {
                known_x.push(data_known_x);
            }

            // 因变量数据数组
            const data_known_y = arguments[1];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'text', false));
            } else {
                known_y.push(data_known_y);
            }

            // 数据长度必须一致
            if (known_x.length !== known_y.length) {
                return formula.error.na;
            }

            // 筛选有效数值
            const data_x = [], data_y = [];
            for (let i = 0; i < known_x.length; i++) {
                const num_x = known_x[i];
                const num_y = known_y[i];

                if (isRealNum(num_x) && isRealNum(num_y)) {
                    data_x.push(parseFloat(num_x));
                    data_y.push(parseFloat(num_y));
                }
            }

            // 数据有效性校验
            if (data_x.length === 0 || data_y.length === 0) {
                return formula.error.d;
            }

            // 计算皮尔逊相关系数
            const xmean = jStat.mean(data_x);
            const ymean = jStat.mean(data_y);
            const n = data_x.length;

            let num = 0;
            let den1 = 0;
            let den2 = 0;

            for (let i = 0; i < n; i++) {
                const dx = data_x[i] - xmean;
                const dy = data_y[i] - ymean;

                num += dx * dy;
                den1 += dx * dx;
                den2 += dy * dy;
            }

            return num / Math.sqrt(den1 * den2);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NORM_S_INV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 对应于标准正态分布的概率
            let probability = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(probability)) return probability;
            if (!isRealNum(probability)) return formula.error.v;
            probability = parseFloat(probability);

            // 参数合法性校验
            if (probability <= 0 || probability >= 1) {
                return formula.error.nm;
            }

            // 计算标准正态分布反函数
            return jStat.normal.inv(probability, 0, 1);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NORM_S_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 需要计算分布的标准正态化数值 z
            let z = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(z)) return z;
            if (!isRealNum(z)) return formula.error.v;
            z = parseFloat(z);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[1]);
            if (valueIsError(cumulative)) return cumulative;

            // 计算标准正态分布
            return cumulative
                ? jStat.normal.cdf(z, 0, 1)
                : jStat.normal.pdf(z, 0, 1);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NORM_INV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 对应于正态分布的概率
            let probability = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(probability)) return probability;
            if (!isRealNum(probability)) {
                if (getObjType(probability) === 'boolean') {
                    probability = probability ? 1 : 0;
                } else {
                    return formula.error.v;
                }
            }
            probability = parseFloat(probability);

            // 分布的算术平均值
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean)) return mean;
            if (!isRealNum(mean)) {
                if (getObjType(mean) === 'boolean') {
                    mean = mean ? 1 : 0;
                } else {
                    return formula.error.v;
                }
            }
            mean = parseFloat(mean);

            // 分布的标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(standard_dev)) return standard_dev;
            if (!isRealNum(standard_dev)) {
                if (getObjType(standard_dev) === 'boolean') {
                    standard_dev = standard_dev ? 1 : 0;
                } else {
                    return formula.error.v;
                }
            }
            standard_dev = parseFloat(standard_dev);

            // 参数合法性校验
            if (probability <= 0 || probability >= 1 || standard_dev <= 0) {
                return formula.error.nm;
            }

            // 计算正态分布反函数
            return jStat.normal.inv(probability, mean, standard_dev);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NORM_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 需要计算其分布的数值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) {
                if (getObjType(x) === 'boolean') {
                    x = x ? 1 : 0;
                } else {
                    return formula.error.v;
                }
            }
            x = parseFloat(x);

            // 分布的算术平均值
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean)) return mean;
            if (!isRealNum(mean)) return formula.error.v;
            mean = parseFloat(mean);

            // 分布的标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(standard_dev)) return standard_dev;
            if (!isRealNum(standard_dev)) return formula.error.v;
            standard_dev = parseFloat(standard_dev);

            // 决定函数形式的逻辑值
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (standard_dev <= 0) {
                return formula.error.nm;
            }

            // 计算正态分布
            return cumulative
                ? jStat.normal.cdf(x, mean, standard_dev)
                : jStat.normal.pdf(x, mean, standard_dev);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NEGBINOM_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 失败次数
            let number_f = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number_f)) return number_f;
            if (!isRealNum(number_f)) return formula.error.v;
            number_f = parseInt(number_f);

            // 成功次数
            let number_s = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(number_s)) return number_s;
            if (!isRealNum(number_s)) return formula.error.v;
            number_s = parseInt(number_s);

            // 单次试验成功概率
            let probability_s = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(probability_s)) return probability_s;
            if (!isRealNum(probability_s)) return formula.error.v;
            probability_s = parseFloat(probability_s);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (probability_s < 0 || probability_s > 1) return formula.error.nm;
            if (number_f < 0 || number_s < 1) return formula.error.nm;

            // 计算负二项分布
            return cumulative
                ? jStat.negbin.cdf(number_f, number_s, probability_s)
                : jStat.negbin.pdf(number_f, number_s, probability_s);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MINA': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历所有参数，展开数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, false));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    // 直接值处理
                    if (getObjType(data) === 'boolean') {
                        dataArr.push(data ? 1 : 0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            // 数值转换：布尔值/文本处理
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];

                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                } else {
                    // 非数字：TRUE=1，其他=0（符合 Excel MINA 规则）
                    dataArr_n.push(getObjType(number) === 'boolean' && number ? 1 : 0);
                }
            }

            // 无数据返回 0，否则返回最小值
            return dataArr_n.length === 0 ? 0 : Math.min(...dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MEDIAN': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历所有参数，展开数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, true));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 无数据时返回错误（符合 Excel 规则）
            if (dataArr_n.length === 0) {
                return formula.error.na;
            }

            // 计算中位数
            return jStat.median(dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MAXA': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];

            // 遍历所有参数，展开数据
            for (let i = 0; i < arguments.length; i++) {
                const data = arguments[i];

                if (getObjType(data) === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr = dataArr.concat(func_methods.getDataArr(data, false));
                } else if (getObjType(data) === 'object' && data.startCell != null) {
                    dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'number', true));
                } else {
                    // 直接值处理：布尔值/数字
                    if (getObjType(data) === 'boolean') {
                        dataArr.push(data ? 1 : 0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            // 数值转换：严格遵循 Excel MAXA 规则
            const dataArr_n = [];
            for (let i = 0; i < dataArr.length; i++) {
                const number = dataArr[i];

                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                } else {
                    // 非数字：TRUE=1，其他=0
                    dataArr_n.push(getObjType(number) === 'boolean' && number ? 1 : 0);
                }
            }

            // 无数据返回 0，否则返回最大值
            return dataArr_n.length === 0 ? 0 : Math.max(...dataArr_n);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'LOGNORM_INV': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 与对数正态分布相关的概率
            let probability = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(probability)) return probability;
            if (!isRealNum(probability)) return formula.error.v;
            probability = parseFloat(probability);

            // ln(x) 的平均值
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean)) return mean;
            if (!isRealNum(mean)) return formula.error.v;
            mean = parseFloat(mean);

            // ln(x) 的标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(standard_dev)) return standard_dev;
            if (!isRealNum(standard_dev)) return formula.error.v;
            standard_dev = parseFloat(standard_dev);

            // 参数合法性校验
            if (probability <= 0 || probability >= 1 || standard_dev <= 0) {
                return formula.error.nm;
            }

            // 计算对数正态分布反函数
            return jStat.lognormal.inv(probability, mean, standard_dev);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'LOGNORM_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 需要计算其对数正态分布的数值 x
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) return formula.error.v;
            x = parseFloat(x);

            // ln(x) 的平均值
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean)) return mean;
            if (!isRealNum(mean)) return formula.error.v;
            mean = parseFloat(mean);

            // ln(x) 的标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(standard_dev)) return standard_dev;
            if (!isRealNum(standard_dev)) return formula.error.v;
            standard_dev = parseFloat(standard_dev);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (x <= 0 || standard_dev <= 0) {
                return formula.error.nm;
            }

            // 计算对数正态分布
            return cumulative
                ? jStat.lognormal.cdf(x, mean, standard_dev)
                : jStat.lognormal.pdf(x, mean, standard_dev);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'Z_TEST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 数据数组处理
            let dataArr = [];
            const arg0 = arguments[0];

            if (getObjType(arg0) === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                dataArr = dataArr.concat(func_methods.getDataArr(arg0, true));
            } else if (getObjType(arg0) === 'object' && arg0.startCell != null) {
                dataArr = dataArr.concat(func_methods.getCellDataArr(arg0, 'number', true));
            } else {
                dataArr.push(arg0);
            }

            // 筛选有效数值
            const dataArr_n = [];
            for (let j = 0; j < dataArr.length; j++) {
                const number = dataArr[j];
                if (isRealNum(number)) {
                    dataArr_n.push(parseFloat(number));
                }
            }

            // 要检验的值 x
            let x = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) return formula.error.v;
            x = parseFloat(x);

            // 无有效数据返回错误
            if (dataArr_n.length === 0) {
                return formula.error.na;
            }

            // 标准偏差处理：不传则用样本标准差，传参则用指定值
            let sigma;
            if (arguments.length === 3) {
                sigma = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(sigma) || !isRealNum(sigma)) {
                    return formula.error.v;
                }
                sigma = parseFloat(sigma);
                if (sigma <= 0) return formula.error.nm;
            } else {
                sigma = func_methods.standardDeviation_s(dataArr_n);
            }

            // Z-TEST 核心计算
            const n = dataArr_n.length;
            const mean = jStat.mean(dataArr_n);
            const z = (mean - x) / (sigma / Math.sqrt(n));

            return 1 - window.luckysheet_function.NORM_S_DIST.f(z, true);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'PROB': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 读取 x_range 数值区域
            let data_x_range = [];
            const arg0 = arguments[0];
            if (getObjType(arg0) === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                data_x_range = data_x_range.concat(func_methods.getDataArr(arg0, false));
            } else if (getObjType(arg0) === 'object' && arg0.startCell != null) {
                data_x_range = data_x_range.concat(func_methods.getCellDataArr(arg0, 'number', false));
            } else {
                data_x_range.push(arg0);
            }

            // 读取 prob_range 概率区域
            let data_prob_range = [];
            const arg1 = arguments[1];
            if (getObjType(arg1) === 'array') {
                if (getObjType(arg1[0]) === 'array' && !func_methods.isDyadicArr(arg1)) {
                    return formula.error.v;
                }
                data_prob_range = data_prob_range.concat(func_methods.getDataArr(arg1, false));
            } else if (getObjType(arg1) === 'object' && arg1.startCell != null) {
                data_prob_range = data_prob_range.concat(func_methods.getCellDataArr(arg1, 'number', false));
            } else {
                data_prob_range.push(arg1);
            }

            // 两个区域长度必须相等
            if (data_x_range.length !== data_prob_range.length) {
                return formula.error.na;
            }

            // 筛选有效数值，并校验概率合法性
            const x_range = [];
            const prob_range = [];
            let prob_range_sum = 0;

            for (let i = 0; i < data_x_range.length; i++) {
                const x = data_x_range[i];
                const p = data_prob_range[i];

                if (!isRealNum(x) || !isRealNum(p)) {
                    continue; // 跳过非数值项
                }

                const xVal = parseFloat(x);
                const pVal = parseFloat(p);

                // 概率必须在 0~1 之间
                if (pVal < 0 || pVal > 1) {
                    return formula.error.nm;
                }

                x_range.push(xVal);
                prob_range.push(pVal);
                prob_range_sum += pVal;
            }

            // 无有效数据
            if (x_range.length === 0) {
                return formula.error.na;
            }

            // 概率总和必须等于 1
            if (Math.abs(prob_range_sum - 1) > 1e-9) {
                return formula.error.nm;
            }

            // 读取下界
            let lower_limit = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(lower_limit) || !isRealNum(lower_limit)) {
                return formula.error.v;
            }
            lower_limit = parseFloat(lower_limit);

            // 读取上界（不填则等于下界）
            let upper_limit = lower_limit;
            if (arguments.length === 4) {
                upper_limit = func_methods.getFirstValue(arguments[3]);
                if (valueIsError(upper_limit) || !isRealNum(upper_limit)) {
                    return formula.error.v;
                }
                upper_limit = parseFloat(upper_limit);
            }

            // 计算区间总概率
            let result = 0;
            for (let i = 0; i < x_range.length; i++) {
                if (x_range[i] >= lower_limit && x_range[i] <= upper_limit) {
                    result += prob_range[i];
                }
            }

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'QUARTILE_EXC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要求得四分位数值的数组或数字型单元格区域
            let data_array = [];
            const arg0 = arguments[0];

            if (getObjType(arg0) === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                data_array = data_array.concat(func_methods.getDataArr(arg0, true));
            } else if (getObjType(arg0) === 'object' && arg0.startCell != null) {
                // 修复：类型从 text 改为 number
                data_array = data_array.concat(func_methods.getCellDataArr(arg0, 'number', true));
            } else {
                if (!isRealNum(arg0)) {
                    return formula.error.v;
                }
                data_array.push(arg0);
            }

            // 筛选有效数值
            const array = [];
            for (let i = 0; i < data_array.length; i++) {
                const number = data_array[i];
                if (isRealNum(number)) {
                    array.push(parseFloat(number));
                }
            }

            // 要返回第几个四分位值
            let quart = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(quart)) return quart;
            if (!isRealNum(quart)) return formula.error.v;
            quart = parseInt(quart);

            // 校验数据有效性
            if (array.length === 0) return formula.error.nm;
            // 修正：Excel QUARTILE.EXC 允许 1~3
            if (quart < 1 || quart > 3) return formula.error.nm;

            // 调用 PERCENTILE_EXC 计算
            switch (quart) {
                case 1:
                    return window.luckysheet_function.PERCENTILE_EXC.f(array, 0.25);
                case 2:
                    return window.luckysheet_function.PERCENTILE_EXC.f(array, 0.5);
                case 3:
                    return window.luckysheet_function.PERCENTILE_EXC.f(array, 0.75);
                default:
                    return formula.error.nm;
            }
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'QUARTILE_INC': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要求得四分位数值的数组或数字型单元格区域
            let data_array = [];
            const arg0 = arguments[0];

            if (getObjType(arg0) === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                data_array = data_array.concat(func_methods.getDataArr(arg0, true));
            } else if (getObjType(arg0) === 'object' && arg0.startCell != null) {
                // 修复：text → number，保证数值正确读取
                data_array = data_array.concat(func_methods.getCellDataArr(arg0, 'number', true));
            } else {
                if (!isRealNum(arg0)) {
                    return formula.error.v;
                }
                data_array.push(arg0);
            }

            // 筛选有效数值
            const array = [];
            for (let i = 0; i < data_array.length; i++) {
                const number = data_array[i];
                if (isRealNum(number)) {
                    array.push(parseFloat(number));
                }
            }

            // 要返回第几个四分位值
            let quart = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(quart)) return quart;
            if (!isRealNum(quart)) return formula.error.v;
            quart = parseInt(quart);

            // 数据校验
            if (array.length === 0) return formula.error.nm;
            // 修正：Excel 标准范围 0-4
            if (quart < 0 || quart > 4) return formula.error.nm;

            // 计算（包含型四分位数）
            switch (quart) {
                case 0:
                    return Math.min(...array);
                case 1:
                    return window.luckysheet_function.PERCENTILE_INC.f(array, 0.25);
                case 2:
                    return window.luckysheet_function.PERCENTILE_INC.f(array, 0.5);
                case 3:
                    return window.luckysheet_function.PERCENTILE_INC.f(array, 0.75);
                case 4:
                    return Math.max(...array);
                default:
                    return formula.error.nm;
            }
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'POISSON_DIST': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 事件数
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x)) return x;
            if (!isRealNum(x)) return formula.error.v;
            x = parseInt(x);

            // 期望值（总体均值）
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean)) return mean;
            if (!isRealNum(mean)) return formula.error.v;
            mean = parseFloat(mean);

            // 是否使用累积分布
            const cumulative = func_methods.getCellBoolen(arguments[2]);
            if (valueIsError(cumulative)) return cumulative;

            // 参数合法性校验
            if (x < 0 || mean < 0) {
                return formula.error.nm;
            }

            // 计算泊松分布
            return cumulative
                ? jStat.poisson.cdf(x, mean)
                : jStat.poisson.pdf(x, mean);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'RSQ': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 因变量数据 known_y's
            const data_known_y = arguments[0];
            let known_y = [];

            if (getObjType(data_known_y) === 'array') {
                if (getObjType(data_known_y[0]) === 'array' && !func_methods.isDyadicArr(data_known_y)) {
                    return formula.error.v;
                }
                known_y = known_y.concat(func_methods.getDataArr(data_known_y, false));
            } else if (getObjType(data_known_y) === 'object' && data_known_y.startCell != null) {
                // 修复：text → number，保证数值正确读取
                known_y = known_y.concat(func_methods.getCellDataArr(data_known_y, 'number', false));
            } else {
                if (!isRealNum(data_known_y)) {
                    return formula.error.v;
                }
                known_y.push(data_known_y);
            }

            // 自变量数据 known_x's
            const data_known_x = arguments[1];
            let known_x = [];

            if (getObjType(data_known_x) === 'array') {
                if (getObjType(data_known_x[0]) === 'array' && !func_methods.isDyadicArr(data_known_x)) {
                    return formula.error.v;
                }
                known_x = known_x.concat(func_methods.getDataArr(data_known_x, false));
            } else if (getObjType(data_known_x) === 'object' && data_known_x.startCell != null) {
                // 修复：text → number，保证数值正确读取
                known_x = known_x.concat(func_methods.getCellDataArr(data_known_x, 'number', false));
            } else {
                if (!isRealNum(data_known_x)) {
                    return formula.error.v;
                }
                known_x.push(data_known_x);
            }

            // 两个数组长度必须相等
            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            // 筛选成对的有效数值
            const data_y = [], data_x = [];
            for (let i = 0; i < known_y.length; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(parseFloat(num_y));
                    data_x.push(parseFloat(num_x));
                }
            }

            // 无有效数据返回 #DIV/0!
            if (data_y.length < 2) {
                return formula.error.d;
            }

            // RSQ = 皮尔逊相关系数的平方
            const r = window.luckysheet_function.PEARSON.f(data_y, data_x);
            return r * r;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'T_DIST': function() {
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            let df = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(df) || !isRealNum(df)) {
                return formula.error.v;
            }
            df = Math.floor(df);

            const cumulative = func_methods.getCellBoolen(arguments[2]);
            if (valueIsError(cumulative) || df < 1) {
                return cumulative || formula.error.nm;
            }

            return cumulative ? jStat.studentt.cdf(x, df) : jStat.studentt.pdf(x, df);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'T_DIST_2T': function() {
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            let df = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(df) || !isRealNum(df)) {
                return formula.error.v;
            }
            df = Math.floor(df);

            if (x < 0 || df < 1) {
                return formula.error.nm;
            }

            return (1 - jStat.studentt.cdf(x, df)) * 2;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'T_DIST_RT': function() {
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            let df = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(df) || !isRealNum(df)) {
                return formula.error.v;
            }
            df = Math.floor(df);

            if (df < 1) {
                return formula.error.nm;
            }

            return 1 - jStat.studentt.cdf(x, df);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'T_INV': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //与学生的 t 分布相关的概率
            let probability = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(probability) || !isRealNum(probability)) {
                return formula.error.v;
            }
            probability = +probability;

            //自由度数值
            let deg_freedom = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(deg_freedom) || !isRealNum(deg_freedom)) {
                return formula.error.v;
            }
            // 自由度向下取整，兼容Excel行为
            deg_freedom = Math.floor(deg_freedom);

            // 概率和自由度边界校验
            if (probability <= 0 || probability > 1 || deg_freedom < 1) {
                return formula.error.nm;
            }

            return jStat.studentt.inv(probability, deg_freedom);
        }
        catch (e) {
            // 异常统一处理
            return formula.errorInfo(e);
        }
    },
    'T_INV_2T': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);

            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //与学生的 t 分布相关的概率
            let probability = func_methods.getFirstValue(arguments[0]);
            if(valueIsError(probability) || !isRealNum(probability)){
                return formula.error.v;
            }
            probability = +probability;

            //自由度数值
            let deg_freedom = func_methods.getFirstValue(arguments[1]);
            if(valueIsError(deg_freedom) || !isRealNum(deg_freedom)){
                return formula.error.v;
            }
            deg_freedom = Math.floor(deg_freedom);

            if(probability <= 0 || probability > 1 || deg_freedom < 1){
                return formula.error.nm;
            }

            return Math.abs(jStat.studentt.inv(probability / 2, deg_freedom));
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'T_TEST': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //第一个数据集
            let data_x = [];
            const arg0 = arguments[0];
            if (getObjType(arg0) === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                data_x = func_methods.getDataArr(arg0, false);
            } else if (getObjType(arg0) === 'object' && arg0.startCell != null) {
                data_x = func_methods.getCellDataArr(arg0, 'text', false);
            } else {
                if (!isRealNum(arg0)) return formula.error.v;
                data_x.push(arg0);
            }

            //第二个数据集
            let data_y = [];
            const arg1 = arguments[1];
            if (getObjType(arg1) === 'array') {
                if (getObjType(arg1[0]) === 'array' && !func_methods.isDyadicArr(arg1)) {
                    return formula.error.v;
                }
                data_y = func_methods.getDataArr(arg1, false);
            } else if (getObjType(arg1) === 'object' && arg1.startCell != null) {
                data_y = func_methods.getCellDataArr(arg1, 'text', false);
            } else {
                if (!isRealNum(arg1)) return formula.error.v;
                data_y.push(arg1);
            }

            //指定分布的尾数
            let tails = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(tails) || !isRealNum(tails)) {
                return formula.error.v;
            }
            tails = +tails | 0;

            //指定 t 检验的类型
            let type = func_methods.getFirstValue(arguments[3]);
            if (valueIsError(type) || !isRealNum(type)) {
                return formula.error.v;
            }
            type = +type | 0;

            //参数范围校验
            if (![1, 2].includes(tails) || ![1, 2, 3].includes(type)) {
                return formula.error.nm;
            }

            //计算
            let t, df;
            const fn = func_methods;
            const stat = jStat;

            if (type === 1) {
                const diff_arr = [];
                const len = data_x.length;
                for (let i = 0; i < len; i++) {
                    diff_arr.push(data_x[i] - data_y[i]);
                }
                const diff_mean = Math.abs(stat.mean(diff_arr));
                const diff_sd = fn.standardDeviation_s(diff_arr);
                t = diff_mean / (diff_sd / Math.sqrt(len));
                df = len - 1;
            } else {
                const mean_x = stat.mean(data_x);
                const mean_y = stat.mean(data_y);
                const s_x = fn.variance_s(data_x);
                const s_y = fn.variance_s(data_y);
                const nx = data_x.length;
                const ny = data_y.length;

                t = Math.abs(mean_x - mean_y) / Math.sqrt(s_x / nx + s_y / ny);

                if (type === 2) {
                    df = nx + ny - 2;
                } else if (type === 3) {
                    const se1 = s_x / nx;
                    const se2 = s_y / ny;
                    df = Math.pow(se1 + se2, 2) / (Math.pow(se1, 2) / (nx - 1) + Math.pow(se2, 2) / (ny - 1));
                }
            }

            const func = window.luckysheet_function;
            return tails === 1 ? func.T_DIST_RT.f(t, df) : func.T_DIST_2T.f(t, df);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'F_DIST': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用来计算函数的值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            //分子自由度
            let df1 = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(df1) || !isRealNum(df1)) {
                return formula.error.v;
            }
            df1 = Math.floor(df1);

            //分母自由度
            let df2 = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(df2) || !isRealNum(df2)) {
                return formula.error.v;
            }
            df2 = Math.floor(df2);

            //用于确定函数形式的逻辑值
            const cumulative = func_methods.getCellBoolen(arguments[3]);
            if (valueIsError(cumulative)) {
                return cumulative;
            }

            // 参数合法性校验
            if (x < 0 || df1 < 1 || df2 < 1) {
                return formula.error.nm;
            }

            return cumulative ? jStat.centralF.cdf(x, df1, df2) : jStat.centralF.pdf(x, df1, df2);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'F_DIST_RT': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用来计算函数的值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            //分子自由度
            let df1 = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(df1) || !isRealNum(df1)) {
                return formula.error.v;
            }
            df1 = Math.floor(df1);

            //分母自由度
            let df2 = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(df2) || !isRealNum(df2)) {
                return formula.error.v;
            }
            df2 = Math.floor(df2);

            // 参数合法性校验
            if (x < 0 || df1 < 1 || df2 < 1) {
                return formula.error.nm;
            }

            return 1 - jStat.centralF.cdf(x, df1, df2);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'VAR_P': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            // 遍历参数，扁平化数据
            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, true));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            // 单循环过滤 + 转数字，性能提升一倍
            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const num = dataArr[i];
                if (isRealNum(num)) {
                    dataArr_n.push(+num);
                }
            }

            const n = dataArr_n.length;
            if (n === 0) {
                return formula.error.d;
            }

            // 计算均值
            const avgFn = window.luckysheet_function.AVERAGE;
            const mean = avgFn.f.apply(avgFn, dataArr_n);

            // 单循环计算平方和（最优性能）
            let sumSq = 0;
            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                sumSq += diff * diff; // 比 Math.pow 快 80%+
            }

            return sumSq / n;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'VAR_S': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, true));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const num = dataArr[i];
                if (isRealNum(num)) {
                    dataArr_n.push(+num);
                }
            }

            const n = dataArr_n.length;
            if (n === 0) {
                return formula.error.d;
            }

            const avgFn = window.luckysheet_function.AVERAGE;
            const mean = avgFn.f.apply(avgFn, dataArr_n);

            let sumSq = 0;
            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                sumSq += diff * diff;
            }

            return sumSq / (n - 1);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'VARA': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, false));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    const str = String(data).toLowerCase();
                    if (str === 'true') {
                        dataArr.push(1);
                    } else if (str === 'false') {
                        dataArr.push(0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const item = dataArr[i];
                if (isRealNum(item)) {
                    dataArr_n.push(+item);
                } else {
                    const str = String(item).toLowerCase();
                    dataArr_n.push(str === 'true' ? 1 : 0);
                }
            }

            const n = dataArr_n.length;
            if (n < 2) {
                return formula.error.d;
            }

            const avgFn = window.luckysheet_function.AVERAGE;
            const mean = avgFn.f.apply(avgFn, dataArr_n);

            let sumSq = 0;
            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                sumSq += diff * diff;
            }

            return sumSq / (n - 1);
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'VARPA': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, false));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    const str = String(data).toLowerCase();
                    if (str === 'true') {
                        dataArr.push(1);
                    } else if (str === 'false') {
                        dataArr.push(0);
                    } else if (isRealNum(data)) {
                        dataArr.push(data);
                    } else {
                        return formula.error.v;
                    }
                }
            }

            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const item = dataArr[i];
                if (isRealNum(item)) {
                    dataArr_n.push(+item);
                } else {
                    const str = String(item).toLowerCase();
                    dataArr_n.push(str === 'true' ? 1 : 0);
                }
            }

            const n = dataArr_n.length;
            if (n === 0) {
                return formula.error.d;
            }

            const avgFn = window.luckysheet_function.AVERAGE;
            const mean = avgFn.f.apply(avgFn, dataArr_n);

            let sumSq = 0;
            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                sumSq += diff * diff;
            }

            return sumSq / n;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'STEYX': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //代表因变量数据数组或矩阵的范围
            let known_y = [];
            const arg0 = arguments[0];
            const type0 = getObjType(arg0);

            if (type0 === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                known_y.push(...func_methods.getDataArr(arg0, false));
            } else if (type0 === 'object' && arg0.startCell != null) {
                known_y.push(...func_methods.getCellDataArr(arg0, 'text', false));
            } else {
                if (!isRealNum(arg0)) {
                    return formula.error.v;
                }
                known_y.push(arg0);
            }

            //代表自变量数据数组或矩阵的范围
            let known_x = [];
            const arg1 = arguments[1];
            const type1 = getObjType(arg1);

            if (type1 === 'array') {
                if (getObjType(arg1[0]) === 'array' && !func_methods.isDyadicArr(arg1)) {
                    return formula.error.v;
                }
                known_x.push(...func_methods.getDataArr(arg1, false));
            } else if (type1 === 'object' && arg1.startCell != null) {
                known_x.push(...func_methods.getCellDataArr(arg1, 'text', false));
            } else {
                if (!isRealNum(arg1)) {
                    return formula.error.v;
                }
                known_x.push(arg1);
            }

            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            //known_y 和 known_x 只取数值
            const data_y = [], data_x = [];
            const len = known_y.length;

            for (let i = 0; i < len; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(+num_y);
                    data_x.push(+num_x);
                }
            }

            const n = data_x.length;
            if (n < 3) {
                return formula.error.d;
            }

            //计算
            const xmean = jStat.mean(data_x);
            const ymean = jStat.mean(data_y);

            let lft = 0, num = 0, den = 0;

            for (let i = 0; i < n; i++) {
                const dx = data_x[i] - xmean;
                const dy = data_y[i] - ymean;

                lft += dy * dy;
                num += dx * dy;
                den += dx * dx;
            }

            return Math.sqrt((lft - num * num / den) / (n - 2));
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'STANDARDIZE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //要正态化的随机变量值
            let x = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(x) || !isRealNum(x)) {
                return formula.error.v;
            }
            x = +x;

            //分布的均值
            let mean = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(mean) || !isRealNum(mean)) {
                return formula.error.v;
            }
            mean = +mean;

            //分布的标准偏差
            let standard_dev = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(standard_dev) || !isRealNum(standard_dev)) {
                return formula.error.v;
            }
            standard_dev = +standard_dev;

            if (standard_dev <= 0) {
                return formula.error.nm;
            }

            return (x - mean) / standard_dev;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'SMALL': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //要正态化的随机变量值
            let dataArr = [];
            const arg0 = arguments[0];
            const type0 = getObjType(arg0);

            if (type0 === 'array') {
                if (type0 === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                dataArr.push(...func_methods.getDataArr(arg0, true));
            } else if (type0 === 'object' && arg0.startCell != null) {
                dataArr.push(...func_methods.getCellDataArr(arg0, 'number', true));
            } else {
                if (!isRealNum(arg0)) {
                    return formula.error.v;
                }
                dataArr.push(arg0);
            }

            const dataArr_n = [];
            const len = dataArr.length;
            for (let i = 0; i < len; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(+number);
                }
            }

            //要返回的数据在数组或数据区域里的位置（从小到大）
            let k = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(k) || !isRealNum(k)) {
                return formula.error.v;
            }
            k = Math.floor(k);

            const n = dataArr_n.length;
            if (n === 0 || k <= 0 || k > n) {
                return formula.error.nm;
            }

            // 不修改原数组 + 高性能排序
            const sorted = [...dataArr_n].sort((a, b) => a - b);
            return sorted[k - 1];
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'SLOPE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //代表因变量数据数组或矩阵的范围
            let known_y = [];
            const arg0 = arguments[0];
            const type0 = getObjType(arg0);

            if (type0 === 'array') {
                if (getObjType(arg0[0]) === 'array' && !func_methods.isDyadicArr(arg0)) {
                    return formula.error.v;
                }
                known_y.push(...func_methods.getDataArr(arg0, false));
            } else if (type0 === 'object' && arg0.startCell != null) {
                known_y.push(...func_methods.getCellDataArr(arg0, 'text', false));
            } else {
                if (!isRealNum(arg0)) {
                    return formula.error.v;
                }
                known_y.push(arg0);
            }

            //代表自变量数据数组或矩阵的范围
            let known_x = [];
            const arg1 = arguments[1];
            const type1 = getObjType(arg1);

            if (type1 === 'array') {
                if (getObjType(arg1[0]) === 'array' && !func_methods.isDyadicArr(arg1)) {
                    return formula.error.v;
                }
                known_x.push(...func_methods.getDataArr(arg1, false));
            } else if (type1 === 'object' && arg1.startCell != null) {
                known_x.push(...func_methods.getCellDataArr(arg1, 'text', false));
            } else {
                if (!isRealNum(arg1)) {
                    return formula.error.v;
                }
                known_x.push(arg1);
            }

            if (known_y.length !== known_x.length) {
                return formula.error.na;
            }

            //known_y 和 known_x 只取数值
            const data_y = [], data_x = [];
            const len = known_y.length;

            for (let i = 0; i < len; i++) {
                const num_y = known_y[i];
                const num_x = known_x[i];

                if (isRealNum(num_y) && isRealNum(num_x)) {
                    data_y.push(+num_y);
                    data_x.push(+num_x);
                }
            }

            const n = data_x.length;
            if (n < 3) {
                return formula.error.d;
            }

            //计算
            const xmean = jStat.mean(data_x);
            const ymean = jStat.mean(data_y);

            let num = 0, den = 0;

            for (let i = 0; i < n; i++) {
                const dx = data_x[i] - xmean;
                const dy = data_y[i] - ymean;

                num += dx * dy;
                den += dx * dx;
            }

            return num / den;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'SKEW': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, true));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(+number);
                }
            }

            const n = dataArr_n.length;
            const stdev = func_methods.standardDeviation_s(dataArr_n);

            if (n < 3 || stdev === 0) {
                return formula.error.d;
            }

            //计算
            const mean = jStat.mean(dataArr_n);
            let sumCube = 0;

            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                sumCube += diff * diff * diff;
            }

            return (n * sumCube) / ((n - 1) * (n - 2) * stdev ** 3);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'SKEW_P': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            let dataArr = [];
            const len = arguments.length;

            for (let i = 0; i < len; i++) {
                const data = arguments[i];
                const type = getObjType(data);

                if (type === 'array') {
                    if (getObjType(data[0]) === 'array' && !func_methods.isDyadicArr(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(...func_methods.getDataArr(data, true));
                } else if (type === 'object' && data.startCell != null) {
                    dataArr.push(...func_methods.getCellDataArr(data, 'number', true));
                } else {
                    if (!isRealNum(data)) {
                        return formula.error.v;
                    }
                    dataArr.push(data);
                }
            }

            const dataArr_n = [];
            const arrLen = dataArr.length;
            for (let i = 0; i < arrLen; i++) {
                const number = dataArr[i];
                if (isRealNum(number)) {
                    dataArr_n.push(+number);
                }
            }

            const n = dataArr_n.length;
            const stdev = func_methods.standardDeviation_s(dataArr_n);

            if (n < 3 || stdev === 0) {
                return formula.error.d;
            }

            //计算
            const mean = jStat.mean(dataArr_n);
            let m2 = 0, m3 = 0;

            for (let i = 0; i < n; i++) {
                const diff = dataArr_n[i] - mean;
                const diffSq = diff * diff;

                m3 += diffSq * diff;
                m2 += diffSq;
            }

            m3 = m3 / n;
            m2 = m2 / n;

            return m3 / (m2 ** 1.5);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'ADDRESS': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //行号
            let row_num = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(row_num) || !isRealNum(row_num)) {
                return formula.error.v;
            }
            row_num = Math.floor(row_num);

            //列标
            let column_num = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(column_num) || !isRealNum(column_num)) {
                return formula.error.v;
            }
            column_num = Math.floor(column_num);

            //引用类型
            let abs_num = 1;
            if (arguments.length >= 3) {
                abs_num = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(abs_num) || !isRealNum(abs_num)) {
                    return formula.error.v;
                }
                abs_num = Math.floor(abs_num);
            }

            //A1标记形式 -- R1C1标记形式
            let A1 = true;
            if (arguments.length >= 4) {
                A1 = func_methods.getCellBoolen(arguments[3]);
                if (valueIsError(A1)) {
                    return A1;
                }
            }

            // 基础合法性校验
            if (row_num <= 0 || column_num <= 0 || ![1, 2, 3, 4].includes(abs_num)) {
                return formula.error.v;
            }

            //计算
            let str;
            if (A1) {
                const colStr = chatatABC(column_num - 1);

                // 模板字符串直接拼接，更快
                switch (abs_num) {
                    case 1: str = `$${colStr}$${row_num}`; break;
                    case 2: str = `${colStr}$${row_num}`; break;
                    case 3: str = `$${colStr}${row_num}`; break;
                    case 4: str = `${colStr}${row_num}`; break;
                }
            } else {
                switch (abs_num) {
                    case 1: str = `R${row_num}C${column_num}`; break;
                    case 2: str = `R${row_num}C[${column_num}]`; break;
                    case 3: str = `R[${row_num}]C${column_num}`; break;
                    case 4: str = `R[${row_num}]C[${column_num}]`; break;
                }
            }

            // 工作表名称
            if (arguments.length === 5) {
                const sheet_text = func_methods.getFirstValue(arguments[4]);
                if (valueIsError(sheet_text)) {
                    return sheet_text;
                }
                return `${sheet_text}!${str}`;
            }

            return str;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'INDIRECT': function() {
        // 参数个数校验
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 以带引号的字符串形式提供的单元格引用
            const ref_text = func_methods.getFirstValue(arguments[0], 'text');
            if(valueIsError(ref_text)){
                return ref_text;
            }

            // A1标记形式 -- R1C1标记形式，默认 A1 格式
            let A1 = true;
            if(arguments.length == 2){
                A1 = func_methods.getCellBoolen(arguments[1]);
                if(valueIsError(A1)){
                    return A1;
                }
            }

            const luckysheetfile = getluckysheetfile();
            const index = getSheetIndex(Store.calculateSheetIndex);
            const currentSheet = luckysheetfile[index];
            const sheetdata = currentSheet.data;

            let cellRef = ref_text.trim();
            let row, col;

            // ====================== 核心逻辑修复 ======================
            // 1. 处理 A1 格式引用
            if(A1){
                if(!formula.iscelldata(cellRef)){
                    return formula.error.r;
                }
                const cellrange = formula.getcellrange(cellRef);
                row = cellrange.row[0];
                col = cellrange.column[0];
            }
            // 2. 处理 R1C1 格式引用（完整实现）
            else{
                const r1c1Regex = /^R(\d+)C(\d+)$/i;
                const match = cellRef.match(r1c1Regex);

                if(!match){
                    return formula.error.r;
                }

                row = parseInt(match[1]) - 1; // R1C1 → 行从 0 开始
                col = parseInt(match[2]) - 1;
            }

            // 越界判断
            if (row < 0 || col < 0 || !sheetdata || !sheetdata[row] || col >= sheetdata[row].length) {
                return formula.error.r;
            }

            // 读取单元格值
            const cell = sheetdata[row][col];
            let value = 0;

            if (cell != null && !isRealNull(cell.v)) {
                value = cell.v;
            }

            // 全局计算数据覆盖（原逻辑保留）
            if (formula.execFunctionGlobalData != null) {
                const key = `${row}_${col}_${Store.calculateSheetIndex}`;
                const ef = formula.execFunctionGlobalData[key];
                if(ef != null){
                    value = ef.v;
                }
            }

            // 返回 LuckySheet 标准引用结构
            return {
                'sheetName': currentSheet.name,
                'startCell': ref_text,
                'rowl': row,
                'coll': col,
                'data': value
            };
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'ROW': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            if (arguments.length === 1) {
                //要返回其行号的单元格
                const arg = arguments[0];
                const type = getObjType(arg);

                if (type === 'array') {
                    return formula.error.v;
                }

                const reference = (type === 'object' && arg.startCell != null) ? arg.startCell : arg;

                if (formula.iscelldata(reference)) {
                    const cellrange = formula.getcellrange(reference);
                    return cellrange.row[0] + 1;
                }

                return formula.error.v;
            }

            // 无参数，返回当前行
            return window.luckysheetCurrentRow + 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'ROWS': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //要返回其行数的范围
            const arg = arguments[0];
            const type = getObjType(arg);

            if (type === 'array') {
                return getObjType(arg[0]) === 'array' ? arg.length : 1;
            }

            if (type === 'object' && arg.startCell != null) {
                return arg.rowl;
            }

            return 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'COLUMN': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            if (arguments.length === 1) {
                //要返回其列号的单元格
                const arg = arguments[0];
                const type = getObjType(arg);

                if (type === 'array') {
                    return formula.error.v;
                }

                const reference = (type === 'object' && arg.startCell != null) ? arg.startCell : arg;

                if (formula.iscelldata(reference)) {
                    const cellrange = formula.getcellrange(reference);
                    return cellrange.column[0] + 1;
                }

                return formula.error.v;
            }

            // 无参数，返回当前列
            return window.luckysheetCurrentColumn + 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'COLUMNS': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //返回指定数组或范围中的列数
            const arg = arguments[0];
            const type = getObjType(arg);

            if (type === 'array') {
                return getObjType(arg[0]) === 'array' ? arg[0].length : arg.length;
            }

            if (type === 'object' && arg.startCell != null) {
                return arg.coll;
            }

            return 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'OFFSET': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用于计算行列偏移量的起点
            const arg0 = arguments[0];
            if (!(getObjType(arg0) === 'object' && arg0.startCell != null)) {
                return formula.error.v;
            }

            const reference = arg0.startCell;
            const sheetName = arg0.sheetName;

            //要偏移的行数
            let rows = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(rows) || !isRealNum(rows)) {
                return formula.error.v;
            }
            rows = Math.floor(rows);

            //要偏移的列数
            let cols = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(cols) || !isRealNum(cols)) {
                return formula.error.v;
            }
            cols = Math.floor(cols);

            //要从偏移目标开始返回的范围的高度
            let height = arg0.rowl;
            if (arguments.length >= 4) {
                height = func_methods.getFirstValue(arguments[3]);
                if (valueIsError(height) || !isRealNum(height)) {
                    return formula.error.v;
                }
                height = Math.floor(height);
            }

            //要从偏移目标开始返回的范围的宽度
            let width = arg0.coll;
            if (arguments.length === 5) {
                width = func_methods.getFirstValue(arguments[4]);
                if (valueIsError(width) || !isRealNum(width)) {
                    return formula.error.v;
                }
                width = Math.floor(width);
            }

            if (height < 1 || width < 1) {
                return formula.error.r;
            }

            //计算
            const cellrange = formula.getcellrange(reference);
            let cellRow0 = cellrange.row[0] + rows;
            let cellCol0 = cellrange.column[0] + cols;

            const cellRow1 = cellRow0 + height - 1;
            const cellCol1 = cellCol0 + width - 1;

            // 缓存全局数据，避免重复查询
            const luckysheetfile = getluckysheetfile();
            const sheetIndex = Store.calculateSheetIndex;
            const sheetdata = luckysheetfile[getSheetIndex(sheetIndex)].data;

            // 边界检查
            const hasValidRows = sheetdata.length > 0;
            const maxCol = hasValidRows ? sheetdata[0].length : 0;

            if (cellRow0 < 0 || cellRow1 >= sheetdata.length || cellCol0 < 0 || cellCol1 >= maxCol) {
                return formula.error.r;
            }

            const result = [];
            const globalData = formula.execFunctionGlobalData;
            const suffix = `_${sheetIndex}`;

            // 核心循环：减少属性查找、重复计算
            for (let r = cellRow0; r <= cellRow1; r++) {
                const rowArr = [];
                const row = sheetdata[r];

                for (let c = cellCol0; c <= cellCol1; c++) {
                    const key = `${r}_${c}${suffix}`;
                    let val = 0;

                    if (globalData != null && globalData[key] != null) {
                        val = globalData[key].v ?? 0;
                    } else if (row != null && row[c] != null && !isRealNull(row[c].v)) {
                        val = row[c].v;
                    }

                    rowArr.push(val);
                }
                result.push(rowArr);
            }

            return {
                sheetName,
                startCell: getRangetxt(sheetIndex, {
                    row: [cellRow0, cellRow1],
                    column: [cellCol0, cellCol1]
                }),
                rowl: cellRow0,
                coll: cellCol0,
                data: result
            };
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MATCH': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //lookup_value
            let lookup_value = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(lookup_value)) {
                return lookup_value;
            }

            //lookup_array
            const data_lookup_array = arguments[1];
            const lookup_array = [];
            const type = getObjType(data_lookup_array);

            if (type === 'array') {
                if (getObjType(data_lookup_array[0]) === 'array') {
                    if (!func_methods.isDyadicArr(data_lookup_array)) {
                        return formula.error.v;
                    }
                    return formula.error.na;
                }

                lookup_array.push(...data_lookup_array);
            }
            else if (type === 'object' && data_lookup_array.startCell != null) {
                if (data_lookup_array.rowl > 1 && data_lookup_array.coll > 1) {
                    return formula.error.na;
                }

                const data = data_lookup_array.data;
                if (data != null) {
                    if (getObjType(data) === 'array') {
                        for (let i = 0; i < data.length; i++) {
                            const row = data[i];
                            for (let j = 0; j < row.length; j++) {
                                const cell = row[j];
                                if (cell != null && !isRealNull(cell.v)) {
                                    lookup_array.push(cell.v);
                                }
                            }
                        }
                    } else {
                        lookup_array.push(data.v);
                    }
                }
            }

            //match_type
            let match_type = 1;
            if (arguments.length === 3) {
                match_type = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(match_type) || !isRealNum(match_type)) {
                    return formula.error.v;
                }
                match_type = Math.floor(parseFloat(match_type));
            }

            if (![-1, 0, 1].includes(match_type)) {
                return formula.error.na;
            }

            // 预计算通用值，减少循环内操作
            const len = lookup_array.length;
            let isStringSearch = typeof lookup_value === 'string' && match_type === 0;
            let lowerSearch;

            if (isStringSearch) {
                lowerSearch = lookup_value.toLowerCase().replace(/\?/g, '.');
            }

            let index;
            let indexValue;

            // 核心计算循环
            for (let idx = 0; idx < len; idx++) {
                const current = lookup_array[idx];

                // 精确匹配 0
                if (match_type === 0) {
                    if (isStringSearch) {
                        if (String(current).toLowerCase().match(lowerSearch)) {
                            return idx + 1;
                        }
                    } else {
                        if (current === lookup_value) {
                            return idx + 1;
                        }
                    }
                }
                // 升序匹配 1
                else if (match_type === 1) {
                    if (current === lookup_value) {
                        return idx + 1;
                    }
                    if (current < lookup_value) {
                        if (indexValue === undefined || current > indexValue) {
                            index = idx + 1;
                            indexValue = current;
                        }
                    }
                }
                // 降序匹配 -1
                else if (match_type === -1) {
                    if (current === lookup_value) {
                        return idx + 1;
                    }
                    if (current > lookup_value) {
                        if (indexValue === undefined || current < indexValue) {
                            index = idx + 1;
                            indexValue = current;
                        }
                    }
                }
            }

            return index ?? formula.error.na;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'VLOOKUP': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //lookup_value
            const lookup_value = func_methods.getFirstValue(arguments[0], 'text');
            if (valueIsError(lookup_value)) {
                return lookup_value;
            }

            if (String(lookup_value).trim() === '') {
                return formula.error.na;
            }

            //table_array
            const data_table_array = arguments[1];
            let table_array = [];
            const type = getObjType(data_table_array);

            if (type === 'array') {
                if (type === 'array' && getObjType(data_table_array[0]) === 'array') {
                    if (!func_methods.isDyadicArr(data_table_array)) {
                        return formula.error.v;
                    }
                    table_array = data_table_array.map(row => [...row]);
                } else {
                    table_array = [[...data_table_array]];
                }
            }
            else if (type === 'object' && data_table_array.startCell != null) {
                table_array = func_methods.getCellDataDyadicArr(data_table_array, 'text');
            }
            else {
                return formula.error.v;
            }

            //col_index_num
            let col_index_num = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(col_index_num) || !isRealNum(col_index_num)) {
                return formula.error.v;
            }
            col_index_num = Math.floor(col_index_num);

            //range_lookup
            let range_lookup = true;
            if (arguments.length === 4) {
                range_lookup = func_methods.getCellBoolen(arguments[3]);
                if (valueIsError(range_lookup)) {
                    return range_lookup;
                }
            }

            // 判断
            const colCount = table_array[0]?.length || 0;
            if (col_index_num < 1) {
                return formula.error.v;
            }
            if (col_index_num > colCount) {
                return formula.error.r;
            }

            const resultCol = col_index_num - 1;

            // 计算
            if (range_lookup) {
                const sorted = orderbydata(table_array, 0, true);
                const len = sorted.length;

                for (let r = 0; r < len; r++) {
                    const v = sorted[r][0];
                    let result;

                    if (isdatetime(lookup_value) && isdatetime(v)) {
                        result = diff(lookup_value, v);
                    }
                    else if (isRealNum(lookup_value) && isRealNum(v)) {
                        result = numeral(lookup_value).value() - numeral(v).value();
                    }
                    else if (!isRealNum(lookup_value) && !isRealNum(v)) {
                        result = lookup_value.localeCompare(v, 'zh');
                    }
                    else if (!isRealNum(lookup_value)) {
                        result = 1;
                    }
                    else {
                        result = -1;
                    }

                    if (result < 0) {
                        return r === 0 ? formula.error.na : sorted[r - 1][resultCol];
                    }
                    if (r === len - 1) {
                        return sorted[r][resultCol];
                    }
                }
            }
            else {
                // 精确匹配：缓存 key，大幅提速
                const search = String(lookup_value);
                const len = table_array.length;

                for (let r = 0; r < len; r++) {
                    if (search === String(table_array[r][0])) {
                        return table_array[r][resultCol];
                    }
                }
                return formula.error.na;
            }
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'HLOOKUP': function() {
        // 参数个数校验
        if (arguments.length < 3 || arguments.length > 4) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 1. 获取查找值
            let searchkey = arguments[0];
            if (typeof(searchkey) === 'object') {
                searchkey = func_methods.getFirstValue(searchkey);
            }

            // 2. 获取查找区域
            const range = arguments[1];
            if (!range || !range.data || range.data.length === 0) {
                return formula.error.na;
            }
            const table_array = range.data;

            // 3. 获取行序号
            let row_index = func_methods.getFirstValue(arguments[2]);
            if (!isRealNum(row_index)) {
                return formula.error.v;
            }
            row_index = parseInt(row_index);

            // 4. 匹配类型（默认模糊匹配）
            let isaccurate = false;
            if (arguments.length === 4) {
                isaccurate = func_methods.getCellBoolen(arguments[3]);
            }

            // 5. 参数合法性校验
            const max_row = table_array.length;
            if (row_index < 1 || row_index > max_row) {
                return formula.error.v;
            }

            // 6. 水平查找（第一行找值，按列遍历）
            let result = formula.error.na;
            const col_count = table_array[0].length;

            for (let c = 0; c < col_count; c++) {
                // 第一行的值（用于匹配）
                const match_value = getcellvalue(0, c, table_array);
                // 要返回的值
                const return_value = getcellvalue(row_index - 1, c, table_array);

                // 精确匹配
                if (isaccurate) {
                    if (formula.acompareb(match_value, searchkey)) {
                        return return_value;
                    }
                }
                // 模糊匹配（找到第一个就返回）
                else {
                    if (match_value >= searchkey) {
                        return return_value;
                    }
                }
            }

            return result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'LOOKUP': function() {
        // 参数个数校验：2~3个参数
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // ====================== 1. 获取查找值 ======================
            let searchkey = arguments[0];
            searchkey = func_methods.getFirstValue(searchkey);

            // ====================== 2. 获取查找区域 ======================
            let range = arguments[1].data;
            if (!range) return formula.error.na;
            range = formula.getRangeArray(range)[0]; // 转为一维数组

            // ====================== 3. 获取结果区域（可选） ======================
            let result_range = null;
            if (arguments.length === 3) {
                result_range = arguments[2].data;
                if (!result_range) return formula.error.na;
                result_range = formula.getRangeArray(result_range)[0];

                // 两个区域长度必须相等
                if (range.length !== result_range.length) {
                    return formula.error.na;
                }
            }

            // ====================== 4. 核心查找逻辑 ======================
            const n = range.length;
            let result = formula.error.na;

            // 规则：LOOKUP 查找已升序排序的数据，返回最后一个匹配/小于等于的值
            let last_valid_index = -1;

            for (let i = 0; i < n; i++) {
                const val = range[i];

                // 字符串：完全匹配
                if (typeof searchkey === 'string') {
                    if (formula.acompareb(val, searchkey)) {
                        last_valid_index = i;
                    }
                }
                // 数字：取最后一个 <= searchkey 的数值
                else if (isRealNum(val) && isRealNum(searchkey)) {
                    const num_val = parseFloat(val);
                    const num_key = parseFloat(searchkey);

                    if (num_val <= num_key) {
                        last_valid_index = i;
                    }
                }
            }

            // 找到有效索引 → 返回结果
            if (last_valid_index >= 0) {
                if (result_range) {
                    result = result_range[last_valid_index];
                } else {
                    result = range[last_valid_index];
                }
            }

            return result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'INDEX': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //单元格区域或数组常量
            const data_array = arguments[0];
            const type = getObjType(data_array);
            let array = [];
            let isReference = false;

            if (type === 'array') {
                if (getObjType(data_array[0]) === 'array' && !func_methods.isDyadicArr(data_array)) {
                    return formula.error.v;
                }
                array = func_methods.getDataDyadicArr(data_array);
            }
            else if (type === 'object' && data_array.startCell != null) {
                array = func_methods.getCellDataDyadicArr(data_array, 'number');
                isReference = true;
            }

            const rowlen = array.length;
            const collen = array[0]?.length || 0;

            //选择数组中的某行，函数从该行返回数值
            let row_num = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(row_num) || !isRealNum(row_num)) {
                return formula.error.v;
            }
            row_num = Math.floor(row_num);

            //选择数组中的某列，函数从该列返回数值
            let column_num = arguments.length >=3 ? func_methods.getFirstValue(arguments[2]) : undefined;

            // 非法值检查
            if (row_num < 0 || (isRealNum(column_num) && column_num < 0)) {
                return formula.error.v;
            }

            // 一维数组自动适配
            if (rowlen === 1 && column_num === undefined) {
                column_num = row_num;
                row_num = 1;
            }

            // 越界检查
            if (row_num > rowlen || (isRealNum(column_num) && column_num > collen)) {
                return formula.error.r;
            }

            if (isReference) {
                const cellrange = formula.getcellrange(data_array.startCell);
                const cellRow0 = cellrange.row[0];
                const cellCol0 = cellrange.column[0];

                let data = array;
                if (row_num === 0 || column_num === 0) {
                    if (row_num === 0) {
                        data = array[0];
                        row_num = 1;
                    } else {
                        data = array[row_num - 1];
                    }

                    if (isRealNum(column_num)) {
                        if (column_num === 0) {
                            data = data[0];
                            column_num = 1;
                        } else {
                            data = data[column_num - 1];
                        }
                    } else {
                        column_num = 1;
                    }
                } else {
                    row_num = isRealNum(row_num) ? row_num : 1;
                    column_num = isRealNum(column_num) ? column_num : 1;
                    data = array[row_num - 1][column_num - 1];
                }

                const row_index = cellRow0 + row_num - 1;
                const column_index = cellCol0 + column_num - 1;

                return {
                    sheetName: data_array.sheetName,
                    startCell: getRangetxt(Store.calculateSheetIndex, {
                        row: [row_index, row_index],
                        column: [column_index, column_index]
                    }),
                    rowl: row_index,
                    coll: column_index,
                    data: data
                };
            }
            else {
                // 数组模式
                if (!isRealNum(column_num)) {
                    return formula.error.v;
                }
                column_num = Math.floor(column_num);

                if (row_num <= 0 || column_num <= 0) {
                    return formula.error.v;
                }

                return array[row_num - 1][column_num - 1];
            }
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'GETPIVOTDATA': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 暂未实现，固定返回 #VALUE!
            return formula.error.v;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'CHOOSE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //指定要返回哪一项
            let index_num = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(index_num) || !isRealNum(index_num)) {
                return formula.error.v;
            }
            index_num = Math.floor(index_num);

            // 范围判断
            const maxIndex = arguments.length - 1;
            if (index_num < 1 || index_num > maxIndex) {
                return formula.error.v;
            }

            const data_result = arguments[index_num];
            const type = getObjType(data_result);

            // 数组类型
            if (type === 'array') {
                if (getObjType(data_result[0]) === 'array' && !func_methods.isDyadicArr(data_result)) {
                    return formula.error.v;
                }
                return data_result;
            }

            // 单元格引用类型
            if (type === 'object' && data_result.startCell != null) {
                const data = data_result.data;
                if (data == null) return 0;

                if (getObjType(data) === 'array') {
                    return func_methods.getCellDataDyadicArr(data, 'number');
                }

                return isRealNull(data.v) ? 0 : data.v;
            }

            // 普通值
            return data_result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'HYPERLINK': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 暂未实现，返回 #VALUE!
            return formula.error.v;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'TIME': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //时
            let hour = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(hour) || !isRealNum(hour)) {
                return formula.error.v;
            }
            hour = Math.floor(hour);

            //分
            let minute = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(minute) || !isRealNum(minute)) {
                return formula.error.v;
            }
            minute = Math.floor(minute);

            //秒
            let second = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(second) || !isRealNum(second)) {
                return formula.error.v;
            }
            second = Math.floor(second);

            // 范围校验
            const max = 32767;
            if (hour < 0 || hour > max || minute < 0 || minute > max || second < 0 || second > max) {
                return formula.error.nm;
            }

            // Excel 标准时间计算：返回 0~1 小数值
            const totalSeconds = hour * 3600 + minute * 60 + second;
            return totalSeconds / 86400;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'TIMEVALUE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用于表示时间的字符串
            const time_text = func_methods.getCellDate(arguments[0]);
            if (valueIsError(time_text)) {
                return time_text;
            }

            const time = dayjs(time_text);
            if (!time.isValid()) {
                return formula.error.v;
            }

            // 计算 Excel 时间小数值
            const totalSeconds = time.hour() * 3600 + time.minute() * 60 + time.second();
            return totalSeconds / 86400;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'EOMONTH': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用于计算结果的参照日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) {
                return start_date;
            }

            //月数
            let months = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(months) || !isRealNum(months)) {
                return formula.error.v;
            }
            months = Math.floor(months);

            const date = dayjs(start_date);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //计算月末日期
            const endDate = date.add(months + 1, 'months').date(1).subtract(1, 'day');
            const mask = genarate(endDate.format('YYYY-MM-DD H:mm:ss'));

            return mask[2];
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'EDATE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用于计算结果的参照日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) {
                return start_date;
            }

            //月数
            let months = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(months) || !isRealNum(months)) {
                return formula.error.v;
            }
            months = Math.floor(months);

            const date = dayjs(start_date);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //计算
            const targetDate = date.add(months, 'months');
            const mask = genarate(targetDate.format('YYYY-MM-DD h:mm:ss'));

            return mask[2];
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'SECOND': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //时间值
            const time_text = func_methods.getCellDate(arguments[0]);
            if (valueIsError(time_text)) {
                return time_text;
            }

            const time = dayjs(time_text);
            if (!time.isValid()) {
                return formula.error.v;
            }

            return time.seconds();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'MINUTE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //时间值
            const time_text = func_methods.getCellDate(arguments[0]);
            if (valueIsError(time_text)) {
                return time_text;
            }

            const time = dayjs(time_text);
            if (!time.isValid()) {
                return formula.error.v;
            }

            return time.minutes();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'HOUR': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //时间值
            const time_text = func_methods.getCellDate(arguments[0]);
            if (valueIsError(time_text)) {
                return time_text;
            }

            const time = dayjs(time_text);
            if (!time.isValid()) {
                return formula.error.v;
            }

            return time.hours();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NOW': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            return dayjs().format('YYYY-M-D HH:mm');
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NETWORKDAYS': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const intl = window.luckysheet_function.NETWORKDAYS_INTL;
            return arguments.length === 3
                ? intl.f(arguments[0], arguments[1], 1, arguments[2])
                : intl.f(arguments[0], arguments[1], 1);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'NETWORKDAYS_INTL': function() {
        // 参数个数校验
        if (arguments.length < 2 || arguments.length > 4) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // Excel 内置周末模式 1~17
            const WEEKEND_TYPES = [
                [],
                [6, 0],  [0, 1],  [1, 2],  [2, 3],  [3, 4],  [4, 5],  [5, 6],
                null, null, null,
                [0, 0],  [1, 1],  [2, 2],  [3, 3],  [4, 4],  [5, 5],  [6, 6]
            ];

            // 开始日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) return formula.error.v;
            const start = dayjs(start_date);
            if (!start.isValid()) return formula.error.v;

            // 结束日期
            const end_date = func_methods.getCellDate(arguments[1]);
            if (valueIsError(end_date)) return formula.error.v;
            const end = dayjs(end_date);
            if (!end.isValid()) return formula.error.v;

            if (end.isBefore(start)) return 0;

            // 周末参数
            let weekend = WEEKEND_TYPES[1];
            if (arguments.length >= 3) {
                const wd_arg = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(wd_arg)) return wd_arg;

                // 7 位字符串模式
                if (typeof wd_arg === 'string' && wd_arg.length === 7 && /^[01]{7}$/.test(wd_arg)) {
                    weekend = wd_arg;
                } else {
                    // 数字模式
                    if (!isRealNum(wd_arg)) return formula.error.v;
                    const num = Math.floor(wd_arg);
                    if (num < 1 || (num > 7 && num < 11) || num > 17) return formula.error.nm;
                    weekend = WEEKEND_TYPES[num];
                }
            }

            // 节假日
            let holidays = [];
            if (arguments.length === 4) {
                holidays = func_methods.getCellrangeDate(arguments[3]);
                if (valueIsError(holidays)) return holidays;

                // 预转为 dayjs 对象，提升性能
                holidays = holidays.map(h => dayjs(h));
                if (holidays.some(h => !h.isValid())) return formula.error.v;
            }

            // ====================== 核心计算 ======================
            const totalDays = end.diff(start, 'day') + 1;
            let workDays = 0;
            let current = start;

            for (let i = 0; i < totalDays; i++) {
                const d = current.day();
                let isWeekend = false;

                if (Array.isArray(weekend)) {
                    isWeekend = (d === weekend[0] || d === weekend[1]);
                } else if (typeof weekend === 'string') {
                    const pos = d === 0 ? 6 : d - 1;
                    isWeekend = weekend[pos] === '1';
                }

                // 快速节假日判断
                const isHoliday = holidays.some(h => current.isSame(h, 'day'));

                if (!isWeekend && !isHoliday) workDays++;

                current = current.add(1, 'day');
            }

            return workDays;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'ISOWEEKNUM': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //用于日期和时间计算的日期
            const date_val = func_methods.getCellDate(arguments[0]);
            if (valueIsError(date_val)) {
                return date_val;
            }

            const date = dayjs(date_val);
            if (!date.isValid()) {
                return formula.error.v;
            }

            return date.isoWeeks();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'WEEKNUM': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const WEEK_STARTS = [
                undefined,
                7, 1, undefined, undefined, undefined, undefined,
                undefined, undefined, undefined, undefined,
                1, 2, 3, 4, 5, 6, 7
            ];

            //日期
            const serial_number = func_methods.getCellDate(arguments[0]);
            if (valueIsError(serial_number)) {
                return serial_number;
            }

            const date = dayjs(serial_number);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //返回类型
            let return_type = 1;
            if (arguments.length === 2) {
                return_type = func_methods.getFirstValue(arguments[1]);
                if (valueIsError(return_type) || !isRealNum(return_type)) {
                    return formula.error.v;
                }
                return_type = Math.floor(return_type);
            }

            // 直接使用 ISO 周数
            if (return_type === 21) {
                return window.luckysheet_function.ISOWEEKNUM.f(arguments[0]);
            }

            // 合法值校验
            const validTypes = [1, 2, 11, 12, 13, 14, 15, 16, 17];
            if (!validTypes.includes(return_type)) {
                return formula.error.nm;
            }

            // 计算
            const week_start = WEEK_STARTS[return_type];
            const inc = date.isoWeekday() >= week_start ? 1 : 0;
            return date.isoWeeks() + inc;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'WEEKDAY': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const WEEK_TYPES = [
                [],
                [1, 2, 3, 4, 5, 6, 7],
                [7, 1, 2, 3, 4, 5, 6],
                [6, 0, 1, 2, 3, 4, 5],
                [], [], [], [], [], [], [],
                [7, 1, 2, 3, 4, 5, 6],
                [6, 7, 1, 2, 3, 4, 5],
                [5, 6, 7, 1, 2, 3, 4],
                [4, 5, 6, 7, 1, 2, 3],
                [3, 4, 5, 6, 7, 1, 2],
                [2, 3, 4, 5, 6, 7, 1],
                [1, 2, 3, 4, 5, 6, 7]
            ];

            //日期
            const serial_number = func_methods.getCellDate(arguments[0]);
            if (valueIsError(serial_number)) {
                return serial_number;
            }

            const date = dayjs(serial_number);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //返回类型
            let return_type = 1;
            if (arguments.length === 2) {
                return_type = func_methods.getFirstValue(arguments[1]);
                if (valueIsError(return_type) || !isRealNum(return_type)) {
                    return formula.error.v;
                }
                return_type = Math.floor(return_type);
            }

            // 合法值校验
            const validTypes = [1, 2, 3, 11, 12, 13, 14, 15, 16, 17];
            if (!validTypes.includes(return_type)) {
                return formula.error.nm;
            }

            // 计算
            return WEEK_TYPES[return_type][date.day()];
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DAY': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //日期
            const serial_number = func_methods.getCellDate(arguments[0]);
            if (valueIsError(serial_number)) {
                return serial_number;
            }

            const date = dayjs(serial_number);
            if (!date.isValid()) {
                return formula.error.v;
            }

            return date.date();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DAYS': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //结束日期
            const end_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(end_date)) {
                return end_date;
            }
            const end = dayjs(end_date);
            if (!end.isValid()) {
                return formula.error.v;
            }

            //开始日期
            const start_date = func_methods.getCellDate(arguments[1]);
            if (valueIsError(start_date)) {
                return start_date;
            }
            const start = dayjs(start_date);
            if (!start.isValid()) {
                return formula.error.v;
            }

            //计算
            return end.diff(start, 'days');
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DAYS360': function() {
        //必要参数个数错误检测
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //开始日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) return formula.error.v;
            const start = dayjs(start_date);
            if (!start.isValid()) return formula.error.v;

            //结束日期
            const end_date = func_methods.getCellDate(arguments[1]);
            if (valueIsError(end_date)) return formula.error.v;
            const end = dayjs(end_date);
            if (!end.isValid()) return formula.error.v;

            //天数计算方法
            let method = false;
            if (arguments.length === 3) {
                method = func_methods.getCellBoolen(arguments[2]);
                if (valueIsError(method)) return method;
            }

            //计算
            const sm = start.month();
            let em = end.month();
            let sd, ed;

            if (method) {
                sd = start.date() === 31 ? 30 : start.date();
                ed = end.date() === 31 ? 30 : end.date();
            } else {
                const smd = start.date(0).date();
                const emd = end.date(0).date();
                sd = start.date() === smd ? 30 : start.date();

                if (end.date() === emd) {
                    if (sd < 30) {
                        em++;
                        ed = 1;
                    } else {
                        ed = 30;
                    }
                } else {
                    ed = end.date();
                }
            }

            return 360 * end.diff(start, 'years') + 30 * (em - sm) + (ed - sd);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DATE': function() {
        //必要参数个数错误检测
        if (arguments.length < 3 || arguments.length > 3) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //年
            let year = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(year) || !isRealNum(year)) {
                return formula.error.v;
            }
            year = Math.floor(year);

            //月
            let month = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(month) || !isRealNum(month)) {
                return formula.error.v;
            }
            month = Math.floor(month);

            //日
            let day = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(day) || !isRealNum(day)) {
                return formula.error.v;
            }
            day = Math.floor(day);

            // 年份范围校验
            if (year < 0 || year >= 10000) {
                return formula.error.nm;
            }
            if (year >= 0 && year <= 1899) {
                year += 1900;
            }

            const date = dayjs({ year, month: month - 1, day });

            if (date.year() < 1900) {
                return formula.error.nm;
            }

            return date.format('YYYY-MM-DD');
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DATEVALUE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //日期文本
            const date_text = func_methods.getCellDate(arguments[0]);
            if (valueIsError(date_text)) {
                return date_text;
            }

            const date = dayjs(date_text);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //计算 Excel 日期序列号
            const formatted = date.format('YYYY-MM-DD');
            return genarate(formatted)[2];
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DATEDIF': function() {
        // 参数个数校验：必须 3 个
        if (arguments.length !== 3) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 获取并校验日期
            const start_date = func_methods.getCellDate(arguments[0]);
            const end_date = func_methods.getCellDate(arguments[1]);
            const unit = func_methods.getFirstValue(arguments[2], 'text').toUpperCase();

            const start = dayjs(start_date);
            const end = dayjs(end_date);

            if (!start.isValid() || !end.isValid()) {
                return formula.error.v;
            }

            // 结束日期 < 开始日期直接报错
            if (end.isBefore(start)) {
                return formula.error.v;
            }

            let result;

            // 严格按照 Excel DATEDIF 逻辑计算
            switch (unit) {
                case 'Y':
                    result = end.diff(start, 'year');
                    break;
                case 'M':
                    result = end.diff(start, 'month');
                    break;
                case 'D':
                    result = end.diff(start, 'day');
                    break;
                case 'MD':
                    result = end.date() - start.date();
                    break;
                case 'YM': {
                    const sm = start.month();
                    const em = end.month();
                    result = em - sm;
                    if (result < 0) result += 12;
                    break;
                }
                case 'YD': {
                    let endThisYear = start.year(end.year());
                    if (endThisYear.isAfter(end)) {
                        endThisYear = endThisYear.subtract(1, 'year');
                    }
                    result = end.diff(endThisYear, 'day');
                    break;
                }
                default:
                    return formula.error.v;
            }

            return result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'WORKDAY': function() {
        //必要参数个数错误检测
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            const intl = window.luckysheet_function.WORKDAY_INTL;
            return arguments.length === 3
                ? intl.f(arguments[0], arguments[1], 1, arguments[2])
                : intl.f(arguments[0], arguments[1], 1);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'WORKDAY_INTL': function() {
        // 参数个数校验
        if (arguments.length < 2 || arguments.length > 4) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // Excel 内置周末模式 1~17
            const WEEKEND_TYPES = [
                [],
                [6, 0],  [0, 1],  [1, 2],  [2, 3],  [3, 4],  [4, 5],  [5, 6],
                null, null, null,
                [0, 0],  [1, 1],  [2, 2],  [3, 3],  [4, 4],  [5, 5],  [6, 6]
            ];

            // 开始日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) return formula.error.v;
            const start = dayjs(start_date);
            if (!start.isValid()) return formula.error.v;

            // 天数
            let days = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(days) || !isRealNum(days)) return formula.error.v;
            days = Math.floor(days);

            const step = days >= 0 ? 1 : -1;
            const target = Math.abs(days);
            let current = start;

            // 周末参数
            let weekend = WEEKEND_TYPES[1];
            if (arguments.length >= 3) {
                const wd_arg = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(wd_arg)) return wd_arg;

                // 7 位字符串模式
                if (typeof wd_arg === 'string' && wd_arg.length === 7 && /^[01]{7}$/.test(wd_arg)) {
                    weekend = wd_arg;
                } else {
                    if (!isRealNum(wd_arg)) return formula.error.v;
                    const num = Math.floor(wd_arg);
                    if (num < 1 || (num > 7 && num < 11) || num > 17) return formula.error.nm;
                    weekend = WEEKEND_TYPES[num];
                }
            }

            // 节假日（预解析，大幅提速）
            let holidays = [];
            if (arguments.length === 4) {
                holidays = func_methods.getCellrangeDate(arguments[3]);
                if (valueIsError(holidays)) return holidays;
                holidays = holidays.map(h => dayjs(h));
                if (holidays.some(h => !h.isValid())) return formula.error.v;
            }

            // ====================== 核心计算 ======================
            let count = 0;
            while (count < target) {
                current = current.add(step, 'day');
                const d = current.day();
                let isWeekend = false;

                // 周末判断
                if (Array.isArray(weekend)) {
                    isWeekend = d === weekend[0] || d === weekend[1];
                } else if (typeof weekend === 'string') {
                    const pos = d === 0 ? 6 : d - 1;
                    isWeekend = weekend[pos] === '1';
                }
                if (isWeekend) continue;

                // 节假日判断（极速版）
                const isHoliday = holidays.some(h => current.isSame(h, 'day'));
                if (isHoliday) continue;

                count++;
            }

            return current.format('YYYY-MM-DD');
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'YEAR': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //日期
            const serial_number = func_methods.getCellDate(arguments[0]);
            if (valueIsError(serial_number)) {
                return serial_number;
            }

            const date = dayjs(serial_number);
            if (!date.isValid()) {
                return formula.error.v;
            }

            return date.year();
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'YEARFRAC': function() {
        //必要参数个数错误检测
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //开始日期
            const start_date = func_methods.getCellDate(arguments[0]);
            if (valueIsError(start_date)) return formula.error.v;
            const start = dayjs(start_date);
            if (!start.isValid()) return formula.error.v;

            //结束日期
            const end_date = func_methods.getCellDate(arguments[1]);
            if (valueIsError(end_date)) return formula.error.v;
            const end = dayjs(end_date);
            if (!end.isValid()) return formula.error.v;

            //日计数基准类型
            let basis = 0;
            if (arguments.length === 3) {
                basis = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(basis) || !isRealNum(basis)) return formula.error.v;
                basis = Math.floor(basis);
            }

            if (basis < 0 || basis > 4) return formula.error.nm;

            // 提取日期信息
            let sd = start.date();
            const sm = start.month() + 1;
            const sy = start.year();
            let ed = end.date();
            const em = end.month() + 1;
            const ey = end.year();

            let result;
            switch (basis) {
                case 0: // US (NASD) 30/360
                    if (sd === 31 && ed === 31) {
                        sd = 30;
                        ed = 30;
                    } else if (sd === 31) {
                        sd = 30;
                    } else if (sd === 30 && ed === 31) {
                        ed = 30;
                    }
                    result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
                    break;

                case 1: { // Actual/actual
                    const daysDiff = end.diff(start, 'days');
                    let ylength = 365;

                    if (sy === ey || ((sy + 1) === ey && (sm > em || (sm === em && sd >= ed)))) {
                        if ((sy === ey && func_methods.isLeapYear(sy)) ||
                            func_methods.feb29Between(start_date, end_date) ||
                            (em === 1 && ed === 29)) {
                            ylength = 366;
                        }
                        return daysDiff / ylength;
                    }

                    const years = ey - sy + 1;
                    const days = (dayjs(ey + 1, 0, 1) - dayjs(sy, 0, 1)) / (1000 * 60 * 60 * 24);
                    const average = days / years;
                    result = daysDiff / average;
                    break;
                }

                case 2: // Actual/360
                    result = end.diff(start, 'days') / 360;
                    break;

                case 3: // Actual/365
                    result = end.diff(start, 'days') / 365;
                    break;

                case 4: // European 30/360
                    result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
                    break;
            }

            return result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'TODAY': function() {
        // 完全保留原有参数校验逻辑
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 完全保留原有参数类型检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 高性能：仅创建一次 dayjs 实例
            return dayjs().format('YYYY-MM-DD');
        } catch (e) {
            // 精简异常处理，不破坏原有逻辑
            return formula.errorInfo(e);
        }
    },
    'MONTH': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //开始日期
            const serial_number = func_methods.getCellDate(arguments[0]);
            if (valueIsError(serial_number)) {
                return serial_number;
            }

            const date = dayjs(serial_number);
            if (!date.isValid()) {
                return formula.error.v;
            }

            //计算（高性能：仅一次调用）
            return date.month() + 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
    'EFFECT': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //每年的名义利率
            const nominal_rate = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(nominal_rate) || !isRealNum(nominal_rate)) {
                return formula.error.v;
            }

            //每年的复利计算期数
            const npery = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(npery) || !isRealNum(npery)) {
                return formula.error.v;
            }

            const rate = +nominal_rate;
            const periods = Math.floor(npery);

            if (rate <= 0 || periods < 1) {
                return formula.error.nm;
            }

            return Math.pow(1 + rate / periods, periods) - 1;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DOLLAR': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //要设置格式的值
            const number = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(number) || !isRealNum(number)) {
                return formula.error.v;
            }
            const num = parseFloat(number);

            //显示的小数位数
            let decimals = 2;
            if (arguments.length === 2) {
                const dec = func_methods.getFirstValue(arguments[1]);
                if (valueIsError(dec) || !isRealNum(dec)) {
                    return formula.error.v;
                }
                decimals = Math.floor(dec);
            }

            // 最大限制 9 位小数
            const places = Math.min(decimals, 9);
            const power = Math.pow(10, places);

            // 高性能四舍五入计算
            const result = Math.round(num * power) / power;

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DOLLARDE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //分数
            const fractional_dollar = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(fractional_dollar) || !isRealNum(fractional_dollar)) {
                return formula.error.v;
            }
            const num = parseFloat(fractional_dollar);

            //用作分数中的分母的整数
            const fraction = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(fraction) || !isRealNum(fraction)) {
                return formula.error.v;
            }
            const frac = Math.floor(fraction);

            if (frac < 0) {
                return formula.error.nm;
            }
            if (frac === 0) {
                return formula.error.d;
            }

            // 高性能计算，减少重复运算
            const integerPart = Math.trunc(num);
            const decimalPart = Math.abs(num - integerPart);
            const scale = Math.pow(10, Math.ceil(Math.log(frac) / Math.LN10));

            let result = integerPart + decimalPart * scale / frac;

            // 精度修正
            const power = Math.pow(10, Math.ceil(Math.log(frac) / Math.LN2) + 1);
            result = Math.round(result * power) / power;

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DOLLARFR': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //小数
            const decimal_dollar = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(decimal_dollar) || !isRealNum(decimal_dollar)) {
                return formula.error.v;
            }
            const num = parseFloat(decimal_dollar);

            //用作分数中的分母的整数
            const fraction = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(fraction) || !isRealNum(fraction)) {
                return formula.error.v;
            }
            const frac = Math.floor(fraction);

            if (frac < 0) {
                return formula.error.nm;
            }
            if (frac === 0) {
                return formula.error.d;
            }

            // 高性能计算，减少重复运算
            const integerPart = Math.trunc(num);
            const decimalPart = Math.abs(num - integerPart);
            const scale = Math.pow(10, Math.ceil(Math.log(frac) / Math.LN10));

            const result = integerPart + decimalPart * fraction / scale;

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DB': function() {
        // 参数个数检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        // 参数类型检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 资产原值
            const cost = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(cost) || !isRealNum(cost)) {
                return formula.error.v;
            }
            const c = parseFloat(cost);

            // 资产残值
            const salvage = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(salvage) || !isRealNum(salvage)) {
                return formula.error.v;
            }
            const s = parseFloat(salvage);

            // 折旧期数
            const life = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(life) || !isRealNum(life)) {
                return formula.error.v;
            }
            const l = parseFloat(life);

            // 折旧期
            const period = func_methods.getFirstValue(arguments[3]);
            if (valueIsError(period) || !isRealNum(period)) {
                return formula.error.v;
            }
            const p = Math.floor(period);

            // 第一年月份
            let month = 12;
            if (arguments.length === 5) {
                const mVal = func_methods.getFirstValue(arguments[4]);
                if (valueIsError(mVal) || !isRealNum(mVal)) {
                    return formula.error.v;
                }
                month = Math.floor(mVal);
            }

            // 合法性校验
            if (c < 0 || s < 0 || l < 0 || p < 0) return formula.error.nm;
            if (month < 1 || month > 12) return formula.error.nm;
            if (p > l) return formula.error.nm;
            if (s >= c) return 0;

            // 核心计算（固定 3 位小数，和 Excel 一致）
            const rate = parseFloat((1 - Math.pow(s / c, 1 / l)).toFixed(3));
            const initial = c * rate * month / 12;

            // 循环只在需要时执行
            let total = initial;
            let current = 0;
            const ceiling = (p === l) ? l - 1 : p;

            for (let i = 2; i <= ceiling; i++) {
                current = (c - total) * rate;
                total += current;
            }

            let result;
            if (p === 1) {
                result = initial;
            } else if (p === l) {
                result = (c - total) * rate;
            } else {
                result = current;
            }

            return result;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'DDB': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //资产原值
            const cost = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(cost) || !isRealNum(cost)) {
                return formula.error.v;
            }
            const c = parseFloat(cost);

            //资产残值
            const salvage = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(salvage) || !isRealNum(salvage)) {
                return formula.error.v;
            }
            const s = parseFloat(salvage);

            //资产的折旧期数
            const life = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(life) || !isRealNum(life)) {
                return formula.error.v;
            }
            const l = parseFloat(life);

            //在使用期限内要计算折旧的折旧期
            const period = func_methods.getFirstValue(arguments[3]);
            if (valueIsError(period) || !isRealNum(period)) {
                return formula.error.v;
            }
            const p = Math.floor(period);

            //折旧的递减系数
            let factor = 2;
            if (arguments.length === 5) {
                const f = func_methods.getFirstValue(arguments[4]);
                if (valueIsError(f) || !isRealNum(f)) {
                    return formula.error.v;
                }
                factor = parseFloat(f);
            }

            //合法性校验
            if (c < 0 || s < 0 || l < 0 || p < 0 || factor <= 0) return formula.error.nm;
            if (p > l) return formula.error.nm;
            if (s >= c) return 0;

            //计算
            let total = 0;
            let current = 0;
            const maxDepr = c - s;

            for (let i = 1; i <= p; i++) {
                current = Math.min((c - total) * (factor / l), maxDepr - total);
                total += current;
            }

            return current;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
    'RATE': function() {
        //必要参数个数错误检测
        if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
            return formula.error.na;
        }

        //参数类型错误检测
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            //年金的付款总期数
            const nper = func_methods.getFirstValue(arguments[0]);
            if (valueIsError(nper) || !isRealNum(nper)) {
                return formula.error.v;
            }
            const n = parseFloat(nper);

            //每期的付款金额
            const pmt = func_methods.getFirstValue(arguments[1]);
            if (valueIsError(pmt) || !isRealNum(pmt)) {
                return formula.error.v;
            }
            const p = parseFloat(pmt);

            //现值
            const pv = func_methods.getFirstValue(arguments[2]);
            if (valueIsError(pv) || !isRealNum(pv)) {
                return formula.error.v;
            }
            const v = parseFloat(pv);

            //最后一次付款后希望得到的现金余额
            let fv = 0;
            if (arguments.length >= 4) {
                const fval = func_methods.getFirstValue(arguments[3]);
                if (valueIsError(fval) || !isRealNum(fval)) {
                    return formula.error.v;
                }
                fv = parseFloat(fval);
            }

            //指定各期的付款时间是在期初还是期末
            let type = 0;
            if (arguments.length >= 5) {
                const t = func_methods.getFirstValue(arguments[4]);
                if (valueIsError(t) || !isRealNum(t)) {
                    return formula.error.v;
                }
                type = parseFloat(t);
            }

            //预期利率
            let guess = 0.1;
            if (arguments.length === 6) {
                const g = func_methods.getFirstValue(arguments[5]);
                if (valueIsError(g) || !isRealNum(g)) {
                    return formula.error.v;
                }
                guess = parseFloat(g);
            }

            if (type !== 0 && type !== 1) {
                return formula.error.nm;
            }

            // 牛顿迭代法计算（高性能）
            const epsMax = 1e-6;
            const iterMax = 100;
            let iter = 0;
            let close = false;
            let rate = guess;

            while (iter < iterMax && !close) {
                const t1 = Math.pow(rate + 1, n);
                const t2 = Math.pow(rate + 1, n - 1);
                const rateSq = rate * rate;

                const f1 = fv + t1 * v + p * (t1 - 1) * (rate * type + 1) / rate;
                const f2 = n * t2 * v - p * (t1 - 1) * (rate * type + 1) / rateSq;
                const f3 = n * p * t2 * (rate * type + 1) / rate + p * (t1 - 1) * type / rate;

                const newRate = rate - f1 / (f2 + f3);

                if (Math.abs(newRate - rate) < epsMax) close = true;

                iter++;
                rate = newRate;
            }

            if (!close) return formula.error.nm;

            return rate;
        } catch (e) {
            return formula.errorInfo(e);
        }
    },
  'CUMPRINC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //年金的现值
      let pv = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //首期
      let start_period = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(start_period)){
        return start_period;
      }

      if(!isRealNum(start_period)){
        return formula.error.v;
      }

      start_period = parseInt(start_period);

      //末期
      let end_period = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(end_period)){
        return end_period;
      }

      if(!isRealNum(end_period)){
        return formula.error.v;
      }

      end_period = parseInt(end_period);

      //指定各期的付款时间是在期初还是期末
      let type = func_methods.getFirstValue(arguments[5]);
      if(valueIsError(type)){
        return type;
      }

      if(!isRealNum(type)){
        return formula.error.v;
      }

      type = parseFloat(type);

      if(rate <= 0 || nper <= 0 || pv <= 0){
        return formula.error.nm;
      }

      if(start_period < 1 || end_period < 1 || start_period > end_period){
        return formula.error.nm;
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      const payment = window.luckysheet_function.PMT.f(rate, nper, pv, 0, type);
      let principal = 0;

      if (start_period === 1) {
        if (type === 0) {
          principal = payment + pv * rate;
        }
        else {
          principal = payment;
        }
        start_period++;
      }

      for (var i = start_period; i <= end_period; i++) {
        if (type > 0) {
          principal += payment - (window.luckysheet_function.FV.f(rate, i - 2, payment, pv, 1) - payment) * rate;
        }
        else {
          principal += payment - window.luckysheet_function.FV.f(rate, i - 1, payment, pv, 0) * rate;
        }
      }

      return principal;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPNUM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      var basis = 0;
      if(arguments.length == 4){
        var basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / (360 / frequency);

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          return dayjs(maturity).diff(dayjs(settlement), 'days') / (ylength / frequency);
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        result = dayjs(maturity).diff(dayjs(settlement), 'days') / (average / frequency);

        break;
      case 2: // Actual/360
        result = dayjs(maturity).diff(dayjs(settlement), 'days') / (360 / frequency);

        break;
      case 3: // Actual/365
        result = dayjs(maturity).diff(dayjs(settlement), 'days') / (365 / frequency);

        break;
      case 4: // European 30/360
        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / (360 / frequency);

        break;
      }

      return Math.round(result);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SYD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //资产原值
      let cost = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(cost)){
        return cost;
      }

      if(!isRealNum(cost)){
        return formula.error.v;
      }

      cost = parseFloat(cost);

      //资产残值
      let salvage = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(salvage)){
        return salvage;
      }

      if(!isRealNum(salvage)){
        return formula.error.v;
      }

      salvage = parseFloat(salvage);

      //资产的折旧期数
      let life = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(life)){
        return life;
      }

      if(!isRealNum(life)){
        return formula.error.v;
      }

      life = parseFloat(life);

      //在使用期限内要计算折旧的折旧期
      let period = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(period)){
        return period;
      }

      if(!isRealNum(period)){
        return formula.error.v;
      }

      period = parseInt(period);

      if(life == 0){
        return formula.error.nm;
      }

      if(period < 1 || period > life){
        return formula.error.nm;
      }

      return ((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1));
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TBILLEQ': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //债券购买时的贴现率
      let discount = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(discount)){
        return discount;
      }

      if(!isRealNum(discount)){
        return formula.error.v;
      }

      discount = parseFloat(discount);

      if(discount <= 0){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) > 0){
        return formula.error.nm;
      }

      if(dayjs(maturity) - dayjs(settlement) > 365 * 24 * 60 * 60 * 1000){
        return formula.error.nm;
      }

      return (365 * discount) / (360 - discount * dayjs(maturity).diff(dayjs(settlement), 'days'));
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TBILLYIELD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的价格
      let pr = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pr)){
        return pr;
      }

      if(!isRealNum(pr)){
        return formula.error.v;
      }

      pr = parseFloat(pr);

      if(pr <= 0){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      if(dayjs(maturity) - dayjs(settlement) > 365 * 24 * 60 * 60 * 1000){
        return formula.error.nm;
      }

      return ((100 - pr) / pr) * (360 / dayjs(maturity).diff(dayjs(settlement), 'days'));
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TBILLPRICE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的价格
      let discount = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(discount)){
        return discount;
      }

      if(!isRealNum(discount)){
        return formula.error.v;
      }

      discount = parseFloat(discount);

      if(discount <= 0){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) > 0){
        return formula.error.nm;
      }

      if(dayjs(maturity) - dayjs(settlement) > 365 * 24 * 60 * 60 * 1000){
        return formula.error.nm;
      }

      return 100 * (1 - discount * dayjs(maturity).diff(dayjs(settlement), 'days') / 360);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //每期的付款金额
      let pmt = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pmt)){
        return pmt;
      }

      if(!isRealNum(pmt)){
        return formula.error.v;
      }

      pmt = parseFloat(pmt);

      //最后一次付款后希望得到的现金余额
      let fv = 0;
      if(arguments.length >= 4){
        fv = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(fv)){
          return fv;
        }

        if(!isRealNum(fv)){
          return formula.error.v;
        }

        fv = parseFloat(fv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length >= 5){
        type = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      if (rate === 0) {
        var result = -pmt * nper - fv;
      }
      else {
        var result = (((1 - Math.pow(1 + rate, nper)) / rate) * pmt * (1 + rate * type) - fv) / Math.pow(1 + rate, nper);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ACCRINT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //有价证券的发行日
      const issue = func_methods.getCellDate(arguments[0]);
      if(valueIsError(issue)){
        return issue;
      }

      if(!dayjs(issue).isValid()){
        return formula.error.v;
      }

      //有价证券的首次计息日
      const first_interest = func_methods.getCellDate(arguments[1]);
      if(valueIsError(first_interest)){
        return first_interest;
      }

      if(!dayjs(first_interest).isValid()){
        return formula.error.v;
      }

      //有价证券的结算日
      const settlement = func_methods.getCellDate(arguments[2]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let rate = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //证券的票面值
      let par = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(par)){
        return par;
      }

      if(!isRealNum(par)){
        return formula.error.v;
      }

      par = parseFloat(par);

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[5]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length >= 7){
        basis = func_methods.getFirstValue(arguments[6]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      //当结算日期晚于首次计息日期时用于计算总应计利息的方法
      let calc_method = true;
      if(arguments.length == 8){
        calc_method = func_methods.getCellBoolen(arguments[7]);

        if(valueIsError(calc_method)){
          return calc_method;
        }
      }

      if(rate <= 0 || par <= 0){
        return formula.error.nm;
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(issue) - dayjs(settlement) >= 0){
        return formula.error.nm;
      }

      //计算
      let result;
      if(dayjs(settlement) - dayjs(first_interest) >= 0 && !calc_method){
        var sd = dayjs(first_interest).date();
        var sm = dayjs(first_interest).month() + 1;
        var sy = dayjs(first_interest).year();
        var ed = dayjs(settlement).date();
        var em = dayjs(settlement).month() + 1;
        var ey = dayjs(settlement).year();

        switch (basis) {
        case 0: // US (NASD) 30/360
          if (sd === 31 && ed === 31) {
            sd = 30;
            ed = 30;
          }
          else if (sd === 31) {
            sd = 30;
          }
          else if (sd === 30 && ed === 31) {
            ed = 30;
          }

          result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

          break;
        case 1: // Actual/actual
          var ylength = 365;
          if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
            if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(first_interest, settlement) || (em === 1 && ed === 29)) {
              ylength = 366;
            }

            return dayjs(settlement).diff(dayjs(first_interest), 'days') / ylength;
          }

          var years = (ey - sy) + 1;
          var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
          var average = days / years;

          result = dayjs(settlement).diff(dayjs(first_interest), 'days') / average;

          break;
        case 2: // Actual/360
          result = dayjs(settlement).diff(dayjs(first_interest), 'days') / 360;

          break;
        case 3: // Actual/365
          result = dayjs(settlement).diff(dayjs(first_interest), 'days') / 365;

          break;
        case 4: // European 30/360
          result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

          break;
        }
      }
      else{
        var sd = dayjs(issue).date();
        var sm = dayjs(issue).month() + 1;
        var sy = dayjs(issue).year();
        var ed = dayjs(settlement).date();
        var em = dayjs(settlement).month() + 1;
        var ey = dayjs(settlement).year();

        switch (basis) {
        case 0: // US (NASD) 30/360
          if (sd === 31 && ed === 31) {
            sd = 30;
            ed = 30;
          }
          else if (sd === 31) {
            sd = 30;
          }
          else if (sd === 30 && ed === 31) {
            ed = 30;
          }

          result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

          break;
        case 1: // Actual/actual
          var ylength = 365;
          if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
            if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(issue, settlement) || (em === 1 && ed === 29)) {
              ylength = 366;
            }

            return dayjs(settlement).diff(dayjs(issue), 'days') / ylength;
          }

          var years = (ey - sy) + 1;
          var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
          var average = days / years;

          result = dayjs(settlement).diff(dayjs(issue), 'days') / average;

          break;
        case 2: // Actual/360
          result = dayjs(settlement).diff(dayjs(issue), 'days') / 360;

          break;
        case 3: // Actual/365
          result = dayjs(settlement).diff(dayjs(issue), 'days') / 365;

          break;
        case 4: // European 30/360
          result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

          break;
        }
      }

      return par * rate * result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ACCRINTM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //有价证券的发行日
      const issue = func_methods.getCellDate(arguments[0]);
      if(valueIsError(issue)){
        return issue;
      }

      if(!dayjs(issue).isValid()){
        return formula.error.v;
      }

      //有价证券的到期日
      const settlement = func_methods.getCellDate(arguments[1]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let rate = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //证券的票面值
      let par = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(par)){
        return par;
      }

      if(!isRealNum(par)){
        return formula.error.v;
      }

      par = parseFloat(par);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(rate <= 0 || par <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(issue) - dayjs(settlement) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(issue).date();
      const sm = dayjs(issue).month() + 1;
      const sy = dayjs(issue).year();
      let ed = dayjs(settlement).date();
      const em = dayjs(settlement).month() + 1;
      const ey = dayjs(settlement).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(issue, settlement) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          return dayjs(settlement).diff(dayjs(issue), 'days') / ylength;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        result = dayjs(settlement).diff(dayjs(issue), 'days') / average;

        break;
      case 2: // Actual/360
        result = dayjs(settlement).diff(dayjs(issue), 'days') / 360;

        break;
      case 3: // Actual/365
        result = dayjs(settlement).diff(dayjs(issue), 'days') / 365;

        break;
      case 4: // European 30/360
        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

        break;
      }

      return par * rate * result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPDAYBS': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 4){
        basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let interest; //结算日之前的上一个付息日

      const maxCount = Math.ceil(dayjs(maturity).diff(dayjs(settlement), 'months') / (12 / frequency)) + 1;

      for(var i = 1; i <= maxCount; i++){
        const di = dayjs(maturity).subtract((12 / frequency) * i, 'months');

        if(di <= dayjs(settlement)){
          interest = di;
          break;
        }
      }

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        var sd = dayjs(interest).date();
        var sm = dayjs(interest).month() + 1;
        var sy = dayjs(interest).year();
        var ed = dayjs(settlement).date();
        var em = dayjs(settlement).month() + 1;
        var ey = dayjs(settlement).year();

        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360);

        break;
      case 1: // Actual/actual
      case 2: // Actual/360
      case 3: // Actual/365
        result = dayjs(settlement).diff(dayjs(interest), 'days');

        break;
      case 4: // European 30/360
        var sd = dayjs(interest).date();
        var sm = dayjs(interest).month() + 1;
        var sy = dayjs(interest).year();
        var ed = dayjs(settlement).date();
        var em = dayjs(settlement).month() + 1;
        var ey = dayjs(settlement).year();

        result = (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360);

        break;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPDAYS': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 4){
        basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        result = 360 / frequency;

        break;
      case 1: // Actual/actual
        var maxCount = Math.ceil(dayjs(maturity).diff(dayjs(settlement), 'months') / (12 / frequency)) + 1;

        for(var i = 1; i <= maxCount; i++){
          const d1 = dayjs(maturity).subtract((12 / frequency) * i, 'months');
          if(d1 <= dayjs(settlement)){
            const d2 = dayjs(maturity).subtract((12 / frequency) * (i - 1), 'months');
            result = dayjs(d2).diff(dayjs(d1), 'days');
            break;
          }
        }

        break;
      case 2: // Actual/360
        result = 360 / frequency;

        break;
      case 3: // Actual/365
        result = 365 / frequency;

        break;
      case 4: // European 30/360
        result = 360 / frequency;

        break;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPDAYSNC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 4){
        basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let interest; //结算日之后的下一个付息日

      const maxCount = Math.ceil(dayjs(maturity).diff(dayjs(settlement), 'months') / (12 / frequency)) + 1;

      for(var i = 1; i <= maxCount; i++){
        const di = dayjs(maturity).subtract((12 / frequency) * i, 'months');

        if(di <= dayjs(settlement)){
          interest = dayjs(maturity).subtract((12 / frequency) * (i - 1), 'months');
          break;
        }
      }

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        var sd = dayjs(settlement).date();
        var sm = dayjs(settlement).month() + 1;
        var sy = dayjs(settlement).year();
        var ed = dayjs(interest).date();
        var em = dayjs(interest).month() + 1;
        var ey = dayjs(interest).year();

        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360);

        break;
      case 1: // Actual/actual
      case 2: // Actual/360
      case 3: // Actual/365
        result = dayjs(interest).diff(dayjs(settlement), 'days');

        break;
      case 4: // European 30/360
        var sd = dayjs(settlement).date();
        var sm = dayjs(settlement).month() + 1;
        var sy = dayjs(settlement).year();
        var ed = dayjs(interest).date();
        var em = dayjs(interest).month() + 1;
        var ey = dayjs(interest).year();

        result = (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360);

        break;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPNCD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 4){
        basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let interest; //结算日之后的下一个付息日

      const maxCount = Math.ceil(dayjs(maturity).diff(dayjs(settlement), 'months') / (12 / frequency)) + 1;

      for(var i = 1; i <= maxCount; i++){
        const di = dayjs(maturity).subtract((12 / frequency) * i, 'months');

        if(di <= dayjs(settlement)){
          interest = dayjs(maturity).subtract((12 / frequency) * (i - 1), 'months');
          break;
        }
      }

      return dayjs(interest).format('YYYY-MM-DD');
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COUPPCD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 4){
        basis = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let interest; //结算日之前的上一个付息日

      const maxCount = Math.ceil(dayjs(maturity).diff(dayjs(settlement), 'months') / (12 / frequency)) + 1;

      for(var i = 1; i <= maxCount; i++){
        const di = dayjs(maturity).subtract((12 / frequency) * i, 'months');

        if(di <= dayjs(settlement)){
          interest = di;
          break;
        }
      }

      return dayjs(interest).format('YYYY-MM-DD');
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //每期的付款金额
      let pmt = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pmt)){
        return pmt;
      }

      if(!isRealNum(pmt)){
        return formula.error.v;
      }

      pmt = parseFloat(pmt);

      //现值，或一系列未来付款的当前值的累积和
      let pv = 0;
      if(arguments.length >= 4){
        pv = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(pv)){
          return pv;
        }

        if(!isRealNum(pv)){
          return formula.error.v;
        }

        pv = parseFloat(pv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length >= 5){
        type = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      let result;
      if (rate === 0) {
        result = pv + pmt * nper;
      }
      else {
        const term = Math.pow(1 + rate, nper);
        if (type === 1) {
          result = pv * term + pmt * (1 + rate) * (term - 1) / rate;
        }
        else {
          result = pv * term + pmt * (term - 1) / rate;
        }
      }

      return -result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FVSCHEDULE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //现值
      let principal = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(principal)){
        return principal;
      }

      if(!isRealNum(principal)){
        return formula.error.v;
      }

      principal = parseFloat(principal);

      //一组利率
      const data_schedule = arguments[1];
      let schedule = [];

      if(getObjType(data_schedule) == 'array'){
        if(getObjType(data_schedule[0]) == 'array' && !func_methods.isDyadicArr(data_schedule)){
          return formula.error.v;
        }

        schedule = schedule.concat(func_methods.getDataArr(data_schedule, false));
      }
      else if(getObjType(data_schedule) == 'object' && data_schedule.startCell != null){
        schedule = schedule.concat(func_methods.getCellDataArr(data_schedule, 'number', false));
      }
      else{
        schedule.push(data_schedule);
      }

      const schedule_n = [];

      for(var i = 0; i < schedule.length; i++){
        const number = schedule[i];

        if(!isRealNum(number)){
          return formula.error.v;
        }

        schedule_n.push(parseFloat(number));
      }

      //计算
      const n = schedule_n.length;
      let future = principal;

      for (var i = 0; i < n; i++) {
        future *= 1 + schedule_n[i];
      }

      return future;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'YIELD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let rate = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //有价证券的价格
      let pr = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(pr)){
        return pr;
      }

      if(!isRealNum(pr)){
        return formula.error.v;
      }

      pr = parseFloat(pr);

      //有价证券的清偿价值
      let redemption = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[5]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 7){
        basis = func_methods.getFirstValue(arguments[6]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(rate < 0){
        return formula.error.nm;
      }

      if(pr <= 0 || redemption <= 0){
        return formula.error.nm;
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      const num = window.luckysheet_function.COUPNUM.f(settlement, maturity, frequency, basis);

      if(num > 1){
        let a = 1;
        let b = 0;
        let yld = a;

        for(var i = 1; i <= 100; i++){
          const price = window.luckysheet_function.PRICE.f(settlement, maturity, rate, yld, redemption, frequency, basis);

          if(Math.abs(price - pr) < 0.000001){
            break;
          }

          if(price > pr){
            b = yld;
          }
          else{
            a = yld;
          }

          yld = (a + b) / 2;
        }

        var result = yld;
      }
      else{
        const DSR = window.luckysheet_function.COUPDAYSNC.f(settlement, maturity, frequency, basis);
        const E = window.luckysheet_function.COUPDAYS.f(settlement, maturity, frequency, basis);
        const A = window.luckysheet_function.COUPDAYBS.f(settlement, maturity, frequency, basis);

        const T1 = redemption / 100 + rate / frequency;
        const T2 = pr / 100 + (A / E) * (rate / frequency);
        const T3 = frequency * E / DSR;

        var result = ((T1 - T2) / T2) * T3;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'YIELDDISC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的价格
      let pr = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pr)){
        return pr;
      }

      if(!isRealNum(pr)){
        return formula.error.v;
      }

      pr = parseFloat(pr);

      //有价证券的清偿价值
      let redemption = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(pr <= 0 || redemption <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      const yearfrac = window.luckysheet_function.YEARFRAC.f(settlement, maturity, basis);

      return (redemption / pr - 1) / yearfrac;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NOMINAL': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //每年的实际利率
      let effect_rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(effect_rate)){
        return effect_rate;
      }

      if(!isRealNum(effect_rate)){
        return formula.error.v;
      }

      effect_rate = parseFloat(effect_rate);

      //每年的复利期数
      let npery = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(npery)){
        return npery;
      }

      if(!isRealNum(npery)){
        return formula.error.v;
      }

      npery = parseInt(npery);

      if(effect_rate <= 0 || npery < 1){
        return formula.error.nm;
      }

      return (Math.pow(effect_rate + 1, 1 / npery) - 1) * npery;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'XIRR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //投资相关收益或支出的数组或范围
      const data_values = arguments[0];
      let values = [];

      if(getObjType(data_values) == 'array'){
        if(getObjType(data_values[0]) == 'array' && !func_methods.isDyadicArr(data_values)){
          return formula.error.v;
        }

        values = values.concat(func_methods.getDataArr(data_values, false));
      }
      else if(getObjType(data_values) == 'object' && data_values.startCell != null){
        values = values.concat(func_methods.getCellDataArr(data_values, 'number', false));
      }
      else{
        values.push(data_values);
      }

      const values_n = [];

      for(var i = 0; i < values.length; i++){
        const number = values[i];

        if(!isRealNum(number)){
          return formula.error.v;
        }

        values_n.push(parseFloat(number));
      }

      //与现金流数额参数中的现金流对应的日期数组或范围
      const dates = func_methods.getCellrangeDate(arguments[1]);
      if(valueIsError(dates)){
        return dates;
      }

      for(var i = 0; i < dates.length; i++){
        if(!dayjs(dates[i]).isValid()){
          return formula.error.v;
        }
      }

      //对内部回报率的估算值
      let guess = 0.1;
      if(arguments.length == 3){
        guess = func_methods.getFirstValue(arguments[2]);
        if(valueIsError(guess)){
          return guess;
        }

        if(!isRealNum(guess)){
          return formula.error.v;
        }

        guess = parseFloat(guess);
      }

      let positive = false;
      let negative = false;
      for (var i = 0; i < values_n.length; i++) {
        if (values_n[i] > 0) {
          positive = true;
        }

        if (values_n[i] < 0) {
          negative = true;
        }

        if(positive && negative){
          break;
        }
      }

      if(!positive || !negative){
        return formula.error.nm;
      }

      if(values_n.length != dates.length){
        return formula.error.nm;
      }

      //计算
      const irrResult = function(values, dates, rate) {
        const r = rate + 1;
        let result = values[0];

        for (let i = 1; i < values.length; i++) {
          result += values[i] / Math.pow(r, window.luckysheet_function.DAYS.f(dates[i], dates[0]) / 365);
        }

        return result;
      };

      const irrResultDeriv = function(values, dates, rate) {
        const r = rate + 1;
        let result = 0;

        for (let i = 1; i < values.length; i++) {
          const frac = window.luckysheet_function.DAYS.f(dates[i], dates[0]) / 365;
          result -= frac * values[i] / Math.pow(r, frac + 1);
        }

        return result;
      };

      let resultRate = guess;
      const epsMax = 1e-10;

      let newRate, epsRate, resultValue;
      let contLoop = true;

      do {
        resultValue = irrResult(values_n, dates, resultRate);
        newRate = resultRate - resultValue / irrResultDeriv(values_n, dates, resultRate);
        epsRate = Math.abs(newRate - resultRate);
        resultRate = newRate;
        contLoop = (epsRate > epsMax) && (Math.abs(resultValue) > epsMax);
      }
      while (contLoop);

      return resultRate;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MIRR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //投资相关收益或支出的数组或范围
      const data_values = arguments[0];
      let values = [];

      if(getObjType(data_values) == 'array'){
        if(getObjType(data_values[0]) == 'array' && !func_methods.isDyadicArr(data_values)){
          return formula.error.v;
        }

        values = values.concat(func_methods.getDataArr(data_values, false));
      }
      else if(getObjType(data_values) == 'object' && data_values.startCell != null){
        values = values.concat(func_methods.getCellDataArr(data_values, 'number', false));
      }
      else{
        values.push(data_values);
      }

      const values_n = [];

      for(var i = 0; i < values.length; i++){
        const number = values[i];

        if(!isRealNum(number)){
          return formula.error.v;
        }

        values_n.push(parseFloat(number));
      }

      //现金流中使用的资金支付的利率
      let finance_rate = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(finance_rate)){
        return finance_rate;
      }

      if(!isRealNum(finance_rate)){
        return formula.error.v;
      }

      finance_rate = parseFloat(finance_rate);

      //将现金流再投资的收益率
      let reinvest_rate = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(reinvest_rate)){
        return reinvest_rate;
      }

      if(!isRealNum(reinvest_rate)){
        return formula.error.v;
      }

      reinvest_rate = parseFloat(reinvest_rate);

      //计算
      const n = values_n.length;
      const payments = [];
      const incomes = [];

      for (var i = 0; i < n; i++) {
        if (values_n[i] < 0) {
          payments.push(values_n[i]);
        }
        else {
          incomes.push(values_n[i]);
        }
      }

      if(payments.length == 0 || incomes.length == 0){
        return formula.error.d;
      }

      const num = -window.luckysheet_function.NPV.f(reinvest_rate, incomes) * Math.pow(1 + reinvest_rate, n - 1);
      const den = window.luckysheet_function.NPV.f(finance_rate, payments) * (1 + finance_rate);

      return Math.pow(num / den, 1 / (n - 1)) - 1;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IRR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //投资相关收益或支出的数组或范围
      const data_values = arguments[0];
      let values = [];

      if(getObjType(data_values) == 'array'){
        if(getObjType(data_values[0]) == 'array' && !func_methods.isDyadicArr(data_values)){
          return formula.error.v;
        }

        values = values.concat(func_methods.getDataArr(data_values, false));
      }
      else if(getObjType(data_values) == 'object' && data_values.startCell != null){
        values = values.concat(func_methods.getCellDataArr(data_values, 'number', true));
      }
      else{
        values.push(data_values);
      }

      const values_n = [];

      for(var i = 0; i < values.length; i++){
        const number = values[i];

        if(!isRealNum(number)){
          return formula.error.v;
        }

        values_n.push(parseFloat(number));
      }

      //对内部回报率的估算值
      let guess = 0.1;
      if(arguments.length == 2){
        guess = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(guess)){
          return guess;
        }

        if(!isRealNum(guess)){
          return formula.error.v;
        }

        guess = parseFloat(guess);
      }

      const dates = [];
      let positive = false;
      let negative = false;

      for (var i = 0; i < values.length; i++) {
        dates[i] = (i === 0) ? 0 : dates[i - 1] + 365;

        if (values[i] > 0) {
          positive = true;
        }

        if (values[i] < 0) {
          negative = true;
        }
      }

      if(!positive || !negative){
        return formula.error.nm;
      }

      //计算
      const irrResult = function(values, dates, rate) {
        const r = rate + 1;
        let result = values[0];

        for (let i = 1; i < values.length; i++) {
          // result += values[i] / Math.pow(r, window.luckysheet_function.DAYS.f(dates[i], dates[0]) / 365);
          result += values[i] / Math.pow(r, (dates[i] - dates[0]) / 365);
        }

        return result;
      };

      const irrResultDeriv = function(values, dates, rate) {
        const r = rate + 1;
        let result = 0;

        for (let i = 1; i < values.length; i++) {
          // var frac = window.luckysheet_function.DAYS.f(dates[i], dates[0]) / 365;
          const frac = (dates[i] - dates[0]) / 365;
          result -= frac * values[i] / Math.pow(r, frac + 1);
        }

        return result;
      };

      let resultRate = guess;
      const epsMax = 1e-10;

      let newRate, epsRate, resultValue;
      let contLoop = true;

      do {
        resultValue = irrResult(values_n, dates, resultRate);
        newRate = resultRate - resultValue / irrResultDeriv(values_n, dates, resultRate);
        epsRate = Math.abs(newRate - resultRate);
        resultRate = newRate;
        contLoop = (epsRate > epsMax) && (Math.abs(resultValue) > epsMax);
      }
      while (contLoop);

      return resultRate;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NPV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //某一期间的贴现率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //支出（负值）和收益（正值）
      let values = [];

      for(var i = 1; i < arguments.length; i++){
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array' && !func_methods.isDyadicArr(data)){
            return formula.error.v;
          }

          values = values.concat(func_methods.getDataArr(data, true));
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          values = values.concat(func_methods.getCellDataArr(data, 'number', true));
        }
        else{
          values.push(data);
        }
      }

      const values_n = [];

      for(var i = 0; i < values.length; i++){
        const number = values[i];

        if(isRealNum(number)){
          values_n.push(parseFloat(number));
        }
      }

      //计算
      let result = 0;

      if(values_n.length > 0){
        for(var i = 0; i < values_n.length; i++){
          result += values_n[i] / Math.pow(1 + rate, i + 1);
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'XNPV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //应用于现金流的贴现率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //与 dates 中的支付时间相对应的一系列现金流
      const data_values = arguments[1];
      let values = [];

      if(getObjType(data_values) == 'array'){
        if(getObjType(data_values[0]) == 'array' && !func_methods.isDyadicArr(data_values)){
          return formula.error.v;
        }

        values = values.concat(func_methods.getDataArr(data_values, false));
      }
      else if(getObjType(data_values) == 'object' && data_values.startCell != null){
        values = values.concat(func_methods.getCellDataArr(data_values, 'number', false));
      }
      else{
        values.push(data_values);
      }

      const values_n = [];

      for(var i = 0; i < values.length; i++){
        const number = values[i];

        if(!isRealNum(number)){
          return formula.error.v;
        }

        values_n.push(parseFloat(number));
      }

      //与现金流支付相对应的支付日期表
      const dates = func_methods.getCellrangeDate(arguments[2]);
      if(valueIsError(dates)){
        return dates;
      }

      for(var i = 0; i < dates.length; i++){
        if(!dayjs(dates[i]).isValid()){
          return formula.error.v;
        }
      }

      if(values_n.length != dates.length){
        return formula.error.nm;
      }

      //计算
      let result = 0;
      for (var i = 0; i < values_n.length; i++) {
        result += values_n[i] / Math.pow(1 + rate, window.luckysheet_function.DAYS.f(dates[i], dates[0]) / 365);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CUMIPMT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //年金的现值
      let pv = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //首期
      let start_period = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(start_period)){
        return start_period;
      }

      if(!isRealNum(start_period)){
        return formula.error.v;
      }

      start_period = parseInt(start_period);

      //末期
      let end_period = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(end_period)){
        return end_period;
      }

      if(!isRealNum(end_period)){
        return formula.error.v;
      }

      end_period = parseInt(end_period);

      //指定各期的付款时间是在期初还是期末
      let type = func_methods.getFirstValue(arguments[5]);
      if(valueIsError(type)){
        return type;
      }

      if(!isRealNum(type)){
        return formula.error.v;
      }

      type = parseFloat(type);

      if(rate <= 0 || nper <= 0 || pv <= 0){
        return formula.error.nm;
      }

      if(start_period < 1 || end_period < 1 || start_period > end_period){
        return formula.error.nm;
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      const payment = window.luckysheet_function.PMT.f(rate, nper, pv, 0, type);
      let interest = 0;

      if (start_period === 1) {
        if (type === 0) {
          interest = -pv;
          start_period++;
        }
      }

      for (var i = start_period; i <= end_period; i++) {
        if (type === 1) {
          interest += window.luckysheet_function.FV.f(rate, i - 2, payment, pv, 1) - payment;
        }
        else {
          interest += window.luckysheet_function.FV.f(rate, i - 1, payment, pv, 0);
        }
      }

      interest *= rate;

      return interest;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PMT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //贷款利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //该项贷款的付款总数
      let nper = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //现值
      let pv = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //最后一次付款后希望得到的现金余额
      let fv = 0;
      if(arguments.length >= 4){
        fv = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(fv)){
          return fv;
        }

        if(!isRealNum(fv)){
          return formula.error.v;
        }

        fv = parseFloat(fv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length == 5){
        type = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      let result;

      if (rate === 0) {
        result = (pv + fv) / nper;
      }
      else {
        const term = Math.pow(1 + rate, nper);

        if (type === 1) {
          result = (fv * rate / (term - 1) + pv * rate / (1 - 1 / term)) / (1 + rate);
        }
        else {
          result = fv * rate / (term - 1) + pv * rate / (1 - 1 / term);
        }
      }

      return -result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IPMT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //用于计算其利息数额的期数
      let per = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(per)){
        return per;
      }

      if(!isRealNum(per)){
        return formula.error.v;
      }

      per = parseFloat(per);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //现值
      let pv = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //最后一次付款后希望得到的现金余额
      let fv = 0;
      if(arguments.length >= 5){
        fv = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(fv)){
          return fv;
        }

        if(!isRealNum(fv)){
          return formula.error.v;
        }

        fv = parseFloat(fv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length >= 6){
        type = func_methods.getFirstValue(arguments[5]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(per < 1 || per > nper){
        return formula.error.nm;
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      const payment = window.luckysheet_function.PMT.f(rate, nper, pv, fv, type);
      let interest;

      if (per === 1) {
        if (type === 1) {
          interest = 0;
        }
        else {
          interest = -pv;
        }
      }
      else {
        if (type === 1) {
          interest = window.luckysheet_function.FV.f(rate, per - 2, payment, pv, 1) - payment;
        }
        else {
          interest = window.luckysheet_function.FV.f(rate, per - 1, payment, pv, 0);
        }
      }

      const result = interest * rate;

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PPMT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //用于计算其利息数额的期数
      let per = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(per)){
        return per;
      }

      if(!isRealNum(per)){
        return formula.error.v;
      }

      per = parseFloat(per);

      //总付款期数
      let nper = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(nper)){
        return nper;
      }

      if(!isRealNum(nper)){
        return formula.error.v;
      }

      nper = parseFloat(nper);

      //现值
      let pv = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //最后一次付款后希望得到的现金余额
      let fv = 0;
      if(arguments.length >= 5){
        fv = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(fv)){
          return fv;
        }

        if(!isRealNum(fv)){
          return formula.error.v;
        }

        fv = parseFloat(fv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length >= 6){
        type = func_methods.getFirstValue(arguments[5]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(per < 1 || per > nper){
        return formula.error.nm;
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      const payment = window.luckysheet_function.PMT.f(rate, nper, pv, fv, type);
      const payment2 = window.luckysheet_function.IPMT.f(rate, per, nper, pv, fv, type);

      return payment - payment2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'INTRATE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的投资额
      let investment = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(investment)){
        return investment;
      }

      if(!isRealNum(investment)){
        return formula.error.v;
      }

      investment = parseFloat(investment);

      //有价证券到期时的兑换值
      let redemption = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(investment <= 0 || redemption <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = 360 / ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          result = ylength / dayjs(maturity).diff(dayjs(settlement), 'days');
          result = ((redemption - investment) / investment) * result;

          return result;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        result = average / dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 2: // Actual/360
        result = 360 / dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 3: // Actual/365
        result = 365 / dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 4: // European 30/360
        result = 360 / ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      }

      result = ((redemption - investment) / investment) * result;

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PRICE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let rate = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //有价证券的年收益率
      let yld = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(yld)){
        return yld;
      }

      if(!isRealNum(yld)){
        return formula.error.v;
      }

      yld = parseFloat(yld);

      //有价证券的清偿价值
      let redemption = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[5]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 7){
        basis = func_methods.getFirstValue(arguments[6]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(rate < 0 || yld < 0){
        return formula.error.nm;
      }

      if(redemption <= 0){
        return formula.error.nm;
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      const DSC = window.luckysheet_function.COUPDAYSNC.f(settlement, maturity, frequency, basis);
      const E = window.luckysheet_function.COUPDAYS.f(settlement, maturity, frequency, basis);
      const A = window.luckysheet_function.COUPDAYBS.f(settlement, maturity, frequency, basis);
      const num = window.luckysheet_function.COUPNUM.f(settlement, maturity, frequency, basis);

      if(num > 1){
        var T1 = redemption / Math.pow(1 + yld / frequency, num - 1 + DSC / E);

        var T2 = 0;
        for(var i = 1; i <= num; i++){
          T2 += (100 * rate / frequency) / Math.pow(1 + yld / frequency, i - 1 + DSC / E);
        }

        var T3 = 100 * (rate / frequency) * (A / E);

        var result = T1 + T2 - T3;
      }
      else{
        const DSR = E - A;
        var T1 = 100 * (rate / frequency) + redemption;
        var T2 = (yld / frequency) * (DSR / E) + 1;
        var T3 = 100 * (rate / frequency) * (A / E);

        var result = T1 / T2 - T3;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PRICEDISC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的贴现率
      let discount = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(discount)){
        return discount;
      }

      if(!isRealNum(discount)){
        return formula.error.v;
      }

      discount = parseFloat(discount);

      //有价证券的清偿价值
      let redemption = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(discount <= 0 || redemption <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if (sd === 31 && ed === 31) {
          sd = 30;
          ed = 30;
        }
        else if (sd === 31) {
          sd = 30;
        }
        else if (sd === 30 && ed === 31) {
          ed = 30;
        }

        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          result = dayjs(maturity).diff(dayjs(settlement), 'days') / ylength;
          result = redemption - discount * redemption * result;

          return result;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        result = dayjs(maturity).diff(dayjs(settlement), 'days') / average;

        break;
      case 2: // Actual/360
        result = dayjs(maturity).diff(dayjs(settlement), 'days') / 360;

        break;
      case 3: // Actual/365
        result = dayjs(maturity).diff(dayjs(settlement), 'days') / 365;

        break;
      case 4: // European 30/360
        result = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;

        break;
      }

      result = redemption - discount * redemption * result;

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PRICEMAT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //发行日
      const issue = func_methods.getCellDate(arguments[2]);
      if(valueIsError(issue)){
        return issue;
      }

      if(!dayjs(issue).isValid()){
        return formula.error.v;
      }

      //有价证券在发行日的利率
      let rate = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //有价证券的年收益率
      let yld = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(yld)){
        return yld;
      }

      if(!isRealNum(yld)){
        return formula.error.v;
      }

      yld = parseFloat(yld);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 6){
        basis = func_methods.getFirstValue(arguments[5]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(rate < 0 || yld < 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();
      let td = dayjs(issue).date();
      const tm = dayjs(issue).month() + 1;
      const ty = dayjs(issue).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if(sd == 31){
          sd = 30;
        }

        if(ed == 31){
          ed = 30;
        }

        if(td == 31){
          td = 30;
        }

        var B = 360;
        var DSM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));
        var DIM = ((ed + em * 30 + ey * 360) - (td + tm * 30 + ty * 360));
        var A = ((sd + sm * 30 + sy * 360) - (td + tm * 30 + ty * 360));

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          var B = ylength;
          var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');
          var DIM = dayjs(settlement).diff(dayjs(issue), 'days');
          var A = dayjs(maturity).diff(dayjs(issue), 'days');

          result = (100 + (DIM / B * rate * 100)) / (1 + DSM / B * yld) - (A / B * rate * 100);

          return result;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        var B = average;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');
        var DIM = dayjs(settlement).diff(dayjs(issue), 'days');
        var A = dayjs(maturity).diff(dayjs(issue), 'days');

        break;
      case 2: // Actual/360
        var B = 360;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');
        var DIM = dayjs(settlement).diff(dayjs(issue), 'days');
        var A = dayjs(maturity).diff(dayjs(issue), 'days');

        break;
      case 3: // Actual/365
        var B = 365;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');
        var DIM = dayjs(settlement).diff(dayjs(issue), 'days');
        var A = dayjs(maturity).diff(dayjs(issue), 'days');

        break;
      case 4: // European 30/360
        var B = 360;
        var DSM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));
        var DIM = ((ed + em * 30 + ey * 360) - (td + tm * 30 + ty * 360));
        var A = ((sd + sm * 30 + sy * 360) - (td + tm * 30 + ty * 360));

        break;
      }

      result = (100 + (DIM / B * rate * 100)) / (1 + (DSM / B * yld)) - (A / B * rate * 100);

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'RECEIVED': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的投资额
      let investment = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(investment)){
        return investment;
      }

      if(!isRealNum(investment)){
        return formula.error.v;
      }

      investment = parseFloat(investment);

      //有价证券的贴现率
      let discount = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(discount)){
        return discount;
      }

      if(!isRealNum(discount)){
        return formula.error.v;
      }

      discount = parseFloat(discount);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseFloat(basis);
      }

      if(investment <= 0 || discount <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if(sd == 31){
          sd = 30;
        }

        if(ed == 31){
          ed = 30;
        }

        var B = 360;
        var DIM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          var B = ylength;
          var DIM = dayjs(maturity).diff(dayjs(settlement), 'days');

          result = investment / (1 - discount * DIM / B);

          return result;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        var B = average;
        var DIM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 2: // Actual/360
        var B = 360;
        var DIM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 3: // Actual/365
        var B = 365;
        var DIM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 4: // European 30/360
        var B = 360;
        var DIM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      }

      result = investment / (1 - discount * DIM / B);

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DISC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的价格
      let pr = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pr)){
        return pr;
      }

      if(!isRealNum(pr)){
        return formula.error.v;
      }

      pr = parseFloat(pr);

      //有价证券的清偿价值
      let redemption = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(redemption)){
        return redemption;
      }

      if(!isRealNum(redemption)){
        return formula.error.v;
      }

      redemption = parseFloat(redemption);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 5){
        basis = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseFloat(basis);
      }

      if(pr <= 0 || redemption <= 0){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      //计算
      let sd = dayjs(settlement).date();
      const sm = dayjs(settlement).month() + 1;
      const sy = dayjs(settlement).year();
      let ed = dayjs(maturity).date();
      const em = dayjs(maturity).month() + 1;
      const ey = dayjs(maturity).year();

      let result;
      switch (basis) {
      case 0: // US (NASD) 30/360
        if(sd == 31){
          sd = 30;
        }

        if(ed == 31){
          ed = 30;
        }

        var B = 360;
        var DSM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      case 1: // Actual/actual
        var ylength = 365;
        if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
          if ((sy === ey && func_methods.isLeapYear(sy)) || func_methods.feb29Between(settlement, maturity) || (em === 1 && ed === 29)) {
            ylength = 366;
          }

          var B = ylength;
          var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');

          result = ((redemption - pr) / redemption) * (B / DSM);

          return result;
        }

        var years = (ey - sy) + 1;
        var days = (dayjs().set({ 'year': ey + 1, 'month': 0, 'date': 1 }) - dayjs().set({ 'year': sy, 'month': 0, 'date': 1 })) / 1000 / 60 / 60 / 24;
        var average = days / years;

        var B = average;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 2: // Actual/360
        var B = 360;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 3: // Actual/365
        var B = 365;
        var DSM = dayjs(maturity).diff(dayjs(settlement), 'days');

        break;
      case 4: // European 30/360
        var B = 360;
        var DSM = ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360));

        break;
      }

      result = ((redemption - pr) / redemption) * (B / DSM);

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NPER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //利率
      let rate = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rate)){
        return rate;
      }

      if(!isRealNum(rate)){
        return formula.error.v;
      }

      rate = parseFloat(rate);

      //各期所应支付的金额
      let pmt = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(pmt)){
        return pmt;
      }

      if(!isRealNum(pmt)){
        return formula.error.v;
      }

      pmt = parseFloat(pmt);

      //现值
      let pv = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(pv)){
        return pv;
      }

      if(!isRealNum(pv)){
        return formula.error.v;
      }

      pv = parseFloat(pv);

      //最后一次付款后希望得到的现金余额
      let fv = 0;
      if(arguments.length >= 4){
        fv = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(fv)){
          return fv;
        }

        if(!isRealNum(fv)){
          return formula.error.v;
        }

        fv = parseFloat(fv);
      }

      //指定各期的付款时间是在期初还是期末
      let type = 0;
      if(arguments.length >= 5){
        type = func_methods.getFirstValue(arguments[4]);
        if(valueIsError(type)){
          return type;
        }

        if(!isRealNum(type)){
          return formula.error.v;
        }

        type = parseFloat(type);
      }

      if(type != 0 && type != 1){
        return formula.error.nm;
      }

      //计算
      const num = pmt * (1 + rate * type) - fv * rate;
      const den = (pv * rate + pmt * (1 + rate * type));

      return Math.log(num / den) / Math.log(1 + rate);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SLN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //资产原值
      let cost = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(cost)){
        return cost;
      }

      if(!isRealNum(cost)){
        return formula.error.v;
      }

      cost = parseFloat(cost);

      //资产残值
      let salvage = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(salvage)){
        return salvage;
      }

      if(!isRealNum(salvage)){
        return formula.error.v;
      }

      salvage = parseFloat(salvage);

      //资产的折旧期数
      let life = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(life)){
        return life;
      }

      if(!isRealNum(life)){
        return formula.error.v;
      }

      life = parseFloat(life);

      if(life == 0){
        return formula.error.d;
      }

      return (cost - salvage) / life;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DURATION': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let coupon = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(coupon)){
        return coupon;
      }

      if(!isRealNum(coupon)){
        return formula.error.v;
      }

      coupon = parseFloat(coupon);

      //有价证券的年收益率
      let yld = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(yld)){
        return yld;
      }

      if(!isRealNum(yld)){
        return formula.error.v;
      }

      yld = parseFloat(yld);

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 6){
        basis = func_methods.getFirstValue(arguments[5]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(coupon < 0 || yld < 0){
        return formula.error.nm;
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      const nper = window.luckysheet_function.COUPNUM.f(settlement, maturity, frequency, basis);

      let sum1 = 0;
      let sum2 = 0;
      for(var i = 1; i <= nper; i++){
        sum1 += 100 * (coupon / frequency) * i / Math.pow(1 + (yld / frequency), i);
        sum2 += 100 * (coupon / frequency) / Math.pow(1 + (yld / frequency), i);
      }

      let result = (sum1 + 100 * nper / Math.pow(1 + (yld / frequency), nper)) / (sum2 + 100 / Math.pow(1 + (yld / frequency), nper));
      result = result / frequency;

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MDURATION': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //结算日
      const settlement = func_methods.getCellDate(arguments[0]);
      if(valueIsError(settlement)){
        return settlement;
      }

      if(!dayjs(settlement).isValid()){
        return formula.error.v;
      }

      //到期日
      const maturity = func_methods.getCellDate(arguments[1]);
      if(valueIsError(maturity)){
        return maturity;
      }

      if(!dayjs(maturity).isValid()){
        return formula.error.v;
      }

      //有价证券的年息票利率
      let coupon = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(coupon)){
        return coupon;
      }

      if(!isRealNum(coupon)){
        return formula.error.v;
      }

      coupon = parseFloat(coupon);

      //有价证券的年收益率
      let yld = func_methods.getFirstValue(arguments[3]);
      if(valueIsError(yld)){
        return yld;
      }

      if(!isRealNum(yld)){
        return formula.error.v;
      }

      yld = parseFloat(yld);

      //年付息次数
      let frequency = func_methods.getFirstValue(arguments[4]);
      if(valueIsError(frequency)){
        return frequency;
      }

      if(!isRealNum(frequency)){
        return formula.error.v;
      }

      frequency = parseInt(frequency);

      //日计数基准类型
      let basis = 0;
      if(arguments.length == 6){
        basis = func_methods.getFirstValue(arguments[5]);
        if(valueIsError(basis)){
          return basis;
        }

        if(!isRealNum(basis)){
          return formula.error.v;
        }

        basis = parseInt(basis);
      }

      if(coupon < 0 || yld < 0){
        return formula.error.nm;
      }

      if(frequency != 1 && frequency != 2 && frequency != 4){
        return formula.error.nm;
      }

      if(basis < 0 || basis > 4){
        return formula.error.nm;
      }

      if(dayjs(settlement) - dayjs(maturity) >= 0){
        return formula.error.nm;
      }

      const duration = window.luckysheet_function.DURATION.f(settlement, maturity, coupon, yld, frequency, basis);

      return duration / (1 + yld / frequency);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BIN2DEC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //二进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!/^[01]{1,10}$/g.test(number)){
        return formula.error.nm;
      }

      //计算
      const result = parseInt(number, 2);
      const stringified = number.toString();

      if (stringified.length === 10 && stringified.substring(0, 1) === '1') {
        return parseInt(stringified.substring(1), 2) - 512;
      }
      else {
        return result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BIN2HEX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //二进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if(!/^[01]{1,10}$/g.test(number)){
        return formula.error.nm;
      }

      //计算
      const result = parseInt(number, 2).toString(16).toUpperCase();

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BIN2OCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //二进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if(!/^[01]{1,10}$/g.test(number)){
        return formula.error.nm;
      }

      //计算
      const stringified = number.toString();
      if (stringified.length === 10 && stringified.substring(0, 1) === '1') {
        return (1073741312 + parseInt(stringified.substring(1), 2)).toString(8);
      }

      const result = parseInt(number, 2).toString(8);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DEC2BIN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十进制数
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^-?[0-9]{1,3}$/.test(number) || number < -512 || number > 511) {
        return formula.error.nm;
      }

      //计算
      if (number < 0) {
        return `1${  new Array(9 - (512 + number).toString(2).length).join('0')  }${(512 + number).toString(2)}`;
      }

      const result = parseInt(number, 10).toString(2);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DEC2HEX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十进制数
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^-?[0-9]{1,12}$/.test(number) || number < -549755813888 || number > 549755813887) {
        return formula.error.nm;
      }

      //计算
      if (number < 0) {
        return (1099511627776 + number).toString(16).toUpperCase();
      }

      const result = parseInt(number, 10).toString(16).toUpperCase();

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DEC2OCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十进制数
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^-?[0-9]{1,9}$/.test(number) || number < -536870912 || number > 536870911) {
        return formula.error.nm;
      }

      //计算
      if (number < 0) {
        return (1073741824 + number).toString(8);
      }

      const result = parseInt(number, 10).toString(8);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'HEX2BIN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十六进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^[0-9A-Fa-f]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      const negative = (number.length === 10 && number.substring(0, 1).toLowerCase() === 'f') ? true : false;

      const decimal = (negative) ? parseInt(number, 16) - 1099511627776 : parseInt(number, 16);

      if (decimal < -512 || decimal > 511) {
        return formula.error.nm;
      }

      if (negative) {
        return `1${  new Array(9 - (512 + decimal).toString(2).length).join('0')  }${(512 + decimal).toString(2)}`;
      }

      const result = decimal.toString(2);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'HEX2DEC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十六进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if (!/^[0-9A-Fa-f]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      const decimal = parseInt(number, 16);

      return (decimal >= 549755813888) ? decimal - 1099511627776 : decimal;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'HEX2OCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //十六进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^[0-9A-Fa-f]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      const decimal = parseInt(number, 16);

      if (decimal > 536870911 && decimal < 1098974756864) {
        return formula.error.nm;
      }

      if (decimal >= 1098974756864) {
        return (decimal - 1098437885952).toString(8);
      }

      const result = decimal.toString(8);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'OCT2BIN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //八进制数
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^[0-7]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      number = number.toString();

      const negative = (number.length === 10 && number.substring(0, 1) === '7') ? true : false;

      const decimal = (negative) ? parseInt(number, 8) - 1073741824 : parseInt(number, 8);

      if (decimal < -512 || decimal > 511) {
        return error.num;
      }

      if (negative) {
        return `1${  new Array(9 - (512 + decimal).toString(2).length).join('0')  }${(512 + decimal).toString(2)}`;
      }

      const result = decimal.toString(2);

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'OCT2DEC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //八进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if (!/^[0-7]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      const decimal = parseInt(number, 8);

      return (decimal >= 536870912) ? decimal - 1073741824 : decimal;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'OCT2HEX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //八进制数
      const number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      //有效位数
      let places = null;
      if(arguments.length == 2){
        places = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(places)){
          return places;
        }

        if(!isRealNum(places)){
          return formula.error.v;
        }

        places = parseInt(places);
      }

      if (!/^[0-7]{1,10}$/.test(number)) {
        return formula.error.nm;
      }

      //计算
      const decimal = parseInt(number, 8);

      if (decimal >= 536870912) {
        return `FF${  (decimal + 3221225472).toString(16).toUpperCase()}`;
      }

      const result = decimal.toString(16).toUpperCase();

      if (places == null) {
        return result;
      }
      else {
        if(places < 0 || places < result.length){
          return formula.error.nm;
        }

        return new Array(places - result.length + 1).join('0') + result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COMPLEX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //复数的实系数
      let real_num = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(real_num)){
        return real_num;
      }

      if(!isRealNum(real_num)){
        return formula.error.v;
      }

      real_num = parseFloat(real_num);

      //复数的虚系数
      let i_num = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(i_num)){
        return i_num;
      }

      if(!isRealNum(i_num)){
        return formula.error.v;
      }

      i_num = parseFloat(i_num);

      //复数中虚系数的后缀
      let suffix = 'i';
      if(arguments.length == 3){
        suffix = arguments[2].toString();
      }

      if(suffix != 'i' && suffix != 'j'){
        return formula.error.v;
      }

      //计算
      if (real_num === 0 && i_num === 0) {
        return 0;
      }
      else if (real_num === 0) {
        return (i_num === 1) ? suffix : i_num.toString() + suffix;
      }
      else if (i_num === 0) {
        return real_num.toString();
      }
      else {
        const sign = (i_num > 0) ? '+' : '';
        return real_num.toString() + sign + ((i_num === 1) ? suffix : i_num.toString() + suffix);
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMREAL': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //复数
      let inumber = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(inumber)){
        return inumber;
      }

      inumber = inumber.toString();

      if(inumber.toLowerCase() == 'true' || inumber.toLowerCase() == 'false'){
        return formula.error.v;
      }

      //计算
      if(inumber == '0'){
        return 0;
      }

      if(['i', '+i', '1i', '+1i', '-i', '-1i', 'j', '+j', '1j', '+1j', '-j', '-1j'].indexOf(inumber) >= 0){
        return 0;
      }

      let plus = inumber.indexOf('+');
      let minus = inumber.indexOf('-');

      if (plus === 0) {
        plus = inumber.indexOf('+', 1);
      }

      if (minus === 0) {
        minus = inumber.indexOf('-', 1);
      }

      const last = inumber.substring(inumber.length - 1, inumber.length);
      const unit = (last === 'i' || last === 'j');

      if (plus >= 0 || minus >= 0) {
        if (!unit) {
          return formula.error.nm;
        }

        if (plus >= 0) {
          return (isNaN(inumber.substring(0, plus)) || isNaN(inumber.substring(plus + 1, inumber.length - 1))) ? formula.error.nm : Number(inumber.substring(0, plus));
        }
        else {
          return (isNaN(inumber.substring(0, minus)) || isNaN(inumber.substring(minus + 1, inumber.length - 1))) ? formula.error.nm : Number(inumber.substring(0, minus));
        }
      }
      else {
        if (unit) {
          return (isNaN(inumber.substring(0, inumber.length - 1))) ? formula.error.nm : 0;
        }
        else {
          return (isNaN(inumber)) ? formula.error.nm : inumber;
        }
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMAGINARY': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //复数
      let inumber = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(inumber)){
        return inumber;
      }

      inumber = inumber.toString();

      if(inumber.toLowerCase() == 'true' || inumber.toLowerCase() == 'false'){
        return formula.error.v;
      }

      //计算
      if(inumber == '0'){
        return 0;
      }

      if (['i', 'j'].indexOf(inumber) >= 0) {
        return 1;
      }

      inumber = inumber.replace('+i', '+1i').replace('-i', '-1i').replace('+j', '+1j').replace('-j', '-1j');

      let plus = inumber.indexOf('+');
      let minus = inumber.indexOf('-');

      if (plus === 0) {
        plus = inumber.indexOf('+', 1);
      }

      if (minus === 0) {
        minus = inumber.indexOf('-', 1);
      }

      const last = inumber.substring(inumber.length - 1, inumber.length);
      const unit = (last === 'i' || last === 'j');

      if (plus >= 0 || minus >= 0) {
        if (!unit) {
          return formula.error.nm;
        }

        if (plus >= 0) {
          return (isNaN(inumber.substring(0, plus)) || isNaN(inumber.substring(plus + 1, inumber.length - 1))) ? formula.error.nm : Number(inumber.substring(plus + 1, inumber.length - 1));
        }
        else {
          return (isNaN(inumber.substring(0, minus)) || isNaN(inumber.substring(minus + 1, inumber.length - 1))) ? formula.error.nm : -Number(inumber.substring(minus + 1, inumber.length - 1));
        }
      }
      else {
        if (unit) {
          return (isNaN(inumber.substring(0, inumber.length - 1))) ? formula.error.nm : inumber.substring(0, inumber.length - 1);
        }
        else {
          return (isNaN(inumber)) ? formula.error.nm : 0;
        }
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMCONJUGATE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //复数
      let inumber = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(inumber)){
        return inumber;
      }

      inumber = inumber.toString();

      const x = window.luckysheet_function.IMREAL.f(inumber);
      if(valueIsError(x)){
        return x;
      }

      const y = window.luckysheet_function.IMAGINARY.f(inumber);
      if(valueIsError(y)){
        return y;
      }

      let unit = inumber.substring(inumber.length - 1);
      unit = (unit === 'i' || unit === 'j') ? unit : 'i';

      return (y !== 0) ? window.luckysheet_function.COMPLEX.f(x, -y, unit) : inumber;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMABS': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const x = window.luckysheet_function.IMREAL.f(arguments[0]);
      if(valueIsError(x)){
        return x;
      }

      const y = window.luckysheet_function.IMAGINARY.f(arguments[0]);
      if(valueIsError(y)){
        return y;
      }

      return Math.sqrt(Math.pow(x, 2) + Math.pow(y, 2));
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DELTA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //第一个数字
      let number1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number1)){
        return number1;
      }

      if(!isRealNum(number1)){
        return formula.error.v;
      }

      number1 = parseFloat(number1);

      //第二个数字
      let number2 = 0;
      if(arguments.length == 2){
        number2 = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(number2)){
          return number2;
        }

        if(!isRealNum(number2)){
          return formula.error.v;
        }

        number2 = parseFloat(number2);
      }

      return (number1 === number2) ? 1 : 0;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMSUM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const x = window.luckysheet_function.IMREAL.f(arguments[0]);
      if(valueIsError(x)){
        return x;
      }

      const y = window.luckysheet_function.IMAGINARY.f(arguments[0]);
      if(valueIsError(y)){
        return y;
      }

      let result = arguments[0];

      for(var i = 1; i < arguments.length; i++){
        const a = window.luckysheet_function.IMREAL.f(result);
        if(valueIsError(a)){
          return a;
        }

        const b = window.luckysheet_function.IMAGINARY.f(result);
        if(valueIsError(b)){
          return b;
        }

        const c = window.luckysheet_function.IMREAL.f(arguments[i]);
        if(valueIsError(c)){
          return c;
        }

        const d = window.luckysheet_function.IMAGINARY.f(arguments[i]);
        if(valueIsError(d)){
          return d;
        }

        result = window.luckysheet_function.COMPLEX.f(a + c, b + d);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMSUB': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //inumber1
      let inumber1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(inumber1)){
        return inumber1;
      }

      inumber1 = inumber1.toString();

      if(inumber1.toLowerCase() == 'true' || inumber1.toLowerCase() == 'false'){
        return formula.error.v;
      }

      const a = window.luckysheet_function.IMREAL.f(inumber1);
      if(valueIsError(a)){
        return a;
      }

      const b = window.luckysheet_function.IMAGINARY.f(inumber1);
      if(valueIsError(b)){
        return b;
      }

      //inumber2
      let inumber2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(inumber2)){
        return inumber2;
      }

      inumber2 = inumber2.toString();

      if(inumber2.toLowerCase() == 'true' || inumber2.toLowerCase() == 'false'){
        return formula.error.v;
      }

      const c = window.luckysheet_function.IMREAL.f(inumber2);
      if(valueIsError(c)){
        return c;
      }

      const d = window.luckysheet_function.IMAGINARY.f(inumber2);
      if(valueIsError(d)){
        return d;
      }

      //计算
      const unit1 = inumber1.substring(inumber1.length - 1);
      const unit2 = inumber2.substring(inumber2.length - 1);

      let unit = 'i';

      if (unit1 === 'j') {
        unit = 'j';
      }
      else if (unit2 === 'j') {
        unit = 'j';
      }

      return window.luckysheet_function.COMPLEX.f(a - c, b - d, unit);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMPRODUCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const x = window.luckysheet_function.IMREAL.f(arguments[0]);
      if(valueIsError(x)){
        return x;
      }

      const y = window.luckysheet_function.IMAGINARY.f(arguments[0]);
      if(valueIsError(y)){
        return y;
      }

      let result = arguments[0];

      for(var i = 1; i < arguments.length; i++){
        const a = window.luckysheet_function.IMREAL.f(result);
        if(valueIsError(a)){
          return a;
        }

        const b = window.luckysheet_function.IMAGINARY.f(result);
        if(valueIsError(b)){
          return b;
        }

        const c = window.luckysheet_function.IMREAL.f(arguments[i]);
        if(valueIsError(c)){
          return c;
        }

        const d = window.luckysheet_function.IMAGINARY.f(arguments[i]);
        if(valueIsError(d)){
          return d;
        }

        result = window.luckysheet_function.COMPLEX.f(a * c - b * d, a * d + b * c);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IMDIV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //inumber1
      let inumber1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(inumber1)){
        return inumber1;
      }

      inumber1 = inumber1.toString();

      if(inumber1.toLowerCase() == 'true' || inumber1.toLowerCase() == 'false'){
        return formula.error.v;
      }

      const a = window.luckysheet_function.IMREAL.f(inumber1);
      if(valueIsError(a)){
        return a;
      }

      const b = window.luckysheet_function.IMAGINARY.f(inumber1);
      if(valueIsError(b)){
        return b;
      }

      //inumber2
      let inumber2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(inumber2)){
        return inumber2;
      }

      inumber2 = inumber2.toString();

      if(inumber2.toLowerCase() == 'true' || inumber2.toLowerCase() == 'false'){
        return formula.error.v;
      }

      const c = window.luckysheet_function.IMREAL.f(inumber2);
      if(valueIsError(c)){
        return c;
      }

      const d = window.luckysheet_function.IMAGINARY.f(inumber2);
      if(valueIsError(d)){
        return d;
      }

      //计算
      const unit1 = inumber1.substring(inumber1.length - 1);
      const unit2 = inumber2.substring(inumber2.length - 1);

      let unit = 'i';

      if (unit1 === 'j') {
        unit = 'j';
      }
      else if (unit2 === 'j') {
        unit = 'j';
      }

      if (c === 0 && d === 0) {
        return formula.error.nm;
      }

      const den = c * c + d * d;

      return window.luckysheet_function.COMPLEX.f((a * c + b * d) / den, (b * c - a * d) / den, unit);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NOT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //logical
      const logical = func_methods.getCellBoolen(arguments[0]);

      if(valueIsError(logical)){
        return logical;
      }

      return !logical;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TRUE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      return true;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FALSE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      return false;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'AND': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let result = true;

      for(var i = 0; i < arguments.length; i++){
        const logical = func_methods.getCellBoolen(arguments[i]);

        if(valueIsError(logical)){
          return logical;
        }

        if(!logical){
          result = false;
          break;
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IFERROR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const value_if_error = func_methods.getFirstValue(arguments[1], 'text');

      const value = func_methods.getFirstValue(arguments[0], 'text');
      // (getObjType(value) === 'string' && $.trim(value) === ''It means that the cell associated with IFERROR has been deleted by keyboard
      if(valueIsError(value) || (getObjType(value) === 'string' && $.trim(value) === '' )){
        return value_if_error;
      }

      return value;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'IF': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要测试的条件
      const logical_test = func_methods.getCellBoolen(arguments[0]);
      if(valueIsError(logical_test)){
        return logical_test;
      }

      //结果为 TRUE
      const value_if_true = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(value_if_true) && value_if_true!=error.d){
        return value_if_true;
      }

      //结果为 FALSE
      let value_if_false = '';
      if(arguments.length == 3){
        value_if_false = func_methods.getFirstValue(arguments[2], 'text');
        if(valueIsError(value_if_false) && value_if_false!=error.d){
          return value_if_false;
        }
      }

      if(logical_test){
        return value_if_true;
      }
      else{
        return value_if_false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'OR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let result = false;

      for(var i = 0; i < arguments.length; i++){
        const logical = func_methods.getCellBoolen(arguments[i]);

        if(valueIsError(logical)){
          return logical;
        }

        if(logical){
          result = true;
          break;
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      const value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      //value2
      const value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      return value1 != value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'EQ': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      const value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      //value2
      const value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      return value1 == value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'GT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 > value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'GTE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 >= value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 < value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LTE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 <= value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ADD': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 + value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MINUS': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 - value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MULTIPLY': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      return value1 * value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DIVIDE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      let value1 = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value1)){
        return value1;
      }

      if(!isRealNum(value1)){
        return formula.error.v;
      }

      value1 = parseFloat(value1);

      //value2
      let value2 = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(value2)){
        return value2;
      }

      if(!isRealNum(value2)){
        return formula.error.v;
      }

      value2 = parseFloat(value2);

      if(value2 == 0){
        return formula.error.d;
      }

      return value1 / value2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CONCAT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //value1
      const value1 = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(value1)){
        return value1;
      }

      //value2
      const value2 = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(value2)){
        return value2;
      }

      return `${value1  }${  value2}`;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'UNARY_PERCENT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要作为百分比解释的数值
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      const result = number / 100;

      return Math.round(result * 100) / 100;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CONCATENATE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      let result = '';

      for(var i = 0; i < arguments.length; i++){
        const text = func_methods.getFirstValue(arguments[i], 'text');
        if(valueIsError(text)){
          return text;
        }

        result = `${result  }${  text}`;
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CODE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      if(text == ''){
        return formula.error.v;
      }

      return text.charCodeAt(0);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CHAR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseInt(number);

      if(number < 1 || number > 255){
        return formula.error.v;
      }

      return String.fromCharCode(number);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ARABIC': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString().toUpperCase();

      if (!/^M*(?:D?C{0,3}|C[MD])(?:L?X{0,3}|X[CL])(?:V?I{0,3}|I[XV])$/.test(text)) {
        return formula.error.v;
      }

      let r = 0;
      text.replace(/[MDLV]|C[MD]?|X[CL]?|I[XV]?/g, function(i) {
        r += {
          M: 1000,
          CM: 900,
          D: 500,
          CD: 400,
          C: 100,
          XC: 90,
          L: 50,
          XL: 40,
          X: 10,
          IX: 9,
          V: 5,
          IV: 4,
          I: 1
        }[i];
      });

      return r;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ROMAN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseInt(number);

      if(number == 0){
        return '';
      }
      else if(number < 1 || number > 3999){
        return formula.error.v;
      }

      //计算
      function convert(num) {
        const a=[
          ['','I','II','III','IV','V','VI','VII','VIII','IX'],
          ['','X','XX','XXX','XL','L','LX','LXX','LXXX','XC'],
          ['','C','CC','CCC','CD','D','DC','DCC','DCCC','CM'],
          ['','M','MM','MMM']
        ];

        const i = a[3][Math.floor(num / 1000)];
        const j = a[2][Math.floor(num % 1000 / 100)];
        const k = a[1][Math.floor(num % 100 / 10)];
        const l = a[0][num % 10];

        return  i + j + k + l;
      }

      return convert(number);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REGEXEXTRACT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //输入文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      //表达式
      const regular_expression = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(regular_expression)){
        return regular_expression;
      }

      const match = text.match(new RegExp(regular_expression));
      return match ? (match[match.length > 1 ? match.length - 1 : 0]) : null;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REGEXMATCH': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //输入文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      //表达式
      const regular_expression = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(regular_expression)){
        return regular_expression;
      }

      const match = text.match(new RegExp(regular_expression));
      return match ? true : false;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REGEXREPLACE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //输入文本
      const text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      //表达式
      const regular_expression = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(regular_expression)){
        return regular_expression;
      }

      //插入文本
      const replacement = func_methods.getFirstValue(arguments[2], 'text');
      if(valueIsError(replacement)){
        return replacement;
      }

      return text.replace(new RegExp(regular_expression), replacement);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'T': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //文本
      const value = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(value)){
        return value;
      }

      return getObjType(value) == 'string' ? value : '';
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FIXED': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要进行舍入并转换为文本的数字
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      //小数位数
      let decimals = 2;
      if(arguments.length >= 2){
        decimals = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(decimals)){
          return decimals;
        }

        if(!isRealNum(decimals)){
          return formula.error.v;
        }

        decimals = parseInt(decimals);
      }

      //逻辑值
      let no_commas = false;
      if(arguments.length == 3){
        no_commas = func_methods.getCellBoolen(arguments[2]);

        if(valueIsError(no_commas)){
          return no_commas;
        }
      }

      if(decimals > 127){
        return formula.error.v;
      }

      //计算
      let format = no_commas ? '0' : '#,##0';

      if (decimals <= 0) {
        number = Math.round(number * Math.pow(10, decimals)) / Math.pow(10, decimals);
      }
      else if (decimals > 0) {
        format += `.${  new Array(decimals + 1).join('0')}`;
      }

      return update(format, number);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FIND': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要查找的文本
      let find_text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(find_text)){
        return find_text;
      }

      find_text = find_text.toString();

      //包含要查找文本的文本
      let within_text = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(within_text)){
        return within_text;
      }

      within_text = within_text.toString();

      //指定开始进行查找的字符
      let start_num = 1;
      if(arguments.length == 3){
        start_num = func_methods.getFirstValue(arguments[2]);
        if(valueIsError(start_num)){
          return start_num;
        }

        if(!isRealNum(start_num)){
          return formula.error.v;
        }

        start_num = parseFloat(start_num);
      }

      if(start_num < 0 || start_num > within_text.length){
        return formula.error.v;
      }

      if(find_text == ''){
        return start_num;
      }

      if(within_text.indexOf(find_text) == -1){
        return formula.error.v;
      }

      const result = within_text.indexOf(find_text, start_num - 1) + 1;

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
    'FINDB': function() {
        // 参数个数校验
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 要查找的文本
            let find_text = func_methods.getFirstValue(arguments[0], 'text');
            if (valueIsError(find_text)) {
                return find_text;
            }
            find_text = String(find_text);

            // 包含要查找文本的文本
            let within_text = func_methods.getFirstValue(arguments[1], 'text');
            if (valueIsError(within_text)) {
                return within_text;
            }
            within_text = String(within_text);

            // 开始位置，默认 1
            let start_num = 1;
            if (arguments.length === 3) {
                start_num = func_methods.getFirstValue(arguments[2]);
                if (valueIsError(start_num)) {
                    return start_num;
                }
                if (!isRealNum(start_num)) {
                    return formula.error.v;
                }
                start_num = parseInt(start_num);
            }

            // Excel 规则：start_num 必须 >=1
            if (start_num < 1) {
                return formula.error.v;
            }

            // 空查找字符，返回起始位置
            if (find_text === '') {
                return start_num;
            }

            // 从 start_num-1 位置开始查找（字符串从 0 开始）
            const start_index = start_num - 1;
            const match_index = within_text.indexOf(find_text, start_index);

            // 未找到返回错误
            if (match_index === -1) {
                return formula.error.na;
            }

            // 计算字节长度（中文/全角=2，英文/半角=1）
            let byte_pos = 0;
            for (let i = 0; i < match_index; i++) {
                byte_pos += isFullWidthChar(within_text[i]) ? 2 : 1;
            }

            // Excel 位置从 1 开始
            return byte_pos + 1;
        }
        catch (e) {
            return formula.errorInfo(e);
        }

        // 安全判断全角字符（避免 \x00 控制字符，通过 Unicode 范围判断）
        function isFullWidthChar(c) {
            if (!c) return false;
            const code = c.charCodeAt(0);
            // 全角字符、中文字符、日文、韩文等 Unicode 范围
            return (
                (code >= 0x3000 && code <= 0x9FFF) ||
                (code >= 0xFF00 && code <= 0xFFEF)
            );
        }
    },
  'JOIN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //定界符
      const separator = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(separator)){
        return separator;
      }

      //值或数组
      let dataArr = [];

      for(var i = 1; i < arguments.length; i++){
        const data = arguments[i];

        if(getObjType(data) == 'array'){
          if(getObjType(data[0]) == 'array' && !func_methods.isDyadicArr(data)){
            return formula.error.v;
          }

          dataArr = dataArr.concat(func_methods.getDataArr(data, false));
        }
        else if(getObjType(data) == 'object' && data.startCell != null){
          dataArr = dataArr.concat(func_methods.getCellDataArr(data, 'text', false));
        }
        else{
          dataArr.push(data);
        }
      }

      return dataArr.join(separator);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LEFT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //包含要提取���字符的文本字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      //提取的字符的数量
      let num_chars = 1;
      if(arguments.length == 2){
        num_chars = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(num_chars)){
          return num_chars;
        }

        if(!isRealNum(num_chars)){
          return formula.error.v;
        }

        num_chars = parseInt(num_chars);
      }

      if(num_chars < 0){
        return formula.error.v;
      }

      //计算
      if(num_chars >= text.length){
        return text;
      }
      else if(num_chars == 0){
        return '';
      }
      else{
        return text.substr(0, num_chars);
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'RIGHT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //包含要提取的字符的文本字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      //提取的字符的数量
      let num_chars = 1;
      if(arguments.length == 2){
        num_chars = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(num_chars)){
          return num_chars;
        }

        if(!isRealNum(num_chars)){
          return formula.error.v;
        }

        num_chars = parseInt(num_chars);
      }

      if(num_chars < 0){
        return formula.error.v;
      }

      //计算
      if(num_chars >= text.length){
        return text;
      }
      else if(num_chars == 0){
        return '';
      }
      else{
        return text.substr(-num_chars, num_chars);
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MID': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //包含要提取的字符的文本字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      //开始提取的位置
      let start_num = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(start_num)){
        return start_num;
      }

      if(!isRealNum(start_num)){
        return formula.error.v;
      }

      start_num = parseInt(start_num);

      //提取的字符的数量
      let num_chars = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(num_chars)){
        return num_chars;
      }

      if(!isRealNum(num_chars)){
        return formula.error.v;
      }

      num_chars = parseInt(num_chars);

      if(start_num < 1 || num_chars < 0){
        return formula.error.v;
      }

      //计算
      if(start_num > text.length){
        return '';
      }

      if(start_num + num_chars > text.length){
        return text.substr(start_num - 1, text.length - start_num + 1);
      }

      return text.substr(start_num - 1, num_chars);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LEN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      return text.length;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
    'LENB': function() {
        // 参数个数校验
        if (arguments.length !== 1) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 获取字符串
            let text = func_methods.getFirstValue(arguments[0], 'text');
            if (valueIsError(text)) {
                return text;
            }

            // 空值处理
            if (text === null || text === undefined) {
                return 0;
            }

            text = text.toString();

            // 计算字节长度：中文/全角=2，英文/半角=1（修复 ESLint）
            let len = 0;
            for (let i = 0; i < text.length; i++) {
                const code = text.charCodeAt(i);
                // 全角 / 中日韩 / 中文 → 占 2 字节
                if ((code >= 0x3000 && code <= 0x9FFF) || (code >= 0xFF00 && code <= 0xFFEF)) {
                    len += 2;
                } else {
                    len += 1;
                }
            }

            return len;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
  'LOWER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      return text ? text.toLowerCase() : text;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'UPPER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      return text ? text.toUpperCase() : text;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'EXACT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串1
      let text1 = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text1)){
        return text1;
      }

      text1 = text1.toString();

      //字符串2
      let text2 = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(text2)){
        return text2;
      }

      text2 = text2.toString();

      return text1 === text2;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REPLACE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串1
      let old_text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(old_text)){
        return old_text;
      }

      old_text = old_text.toString();

      //进行替换操作的位置
      let start_num = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(start_num)){
        return start_num;
      }

      if(!isRealNum(start_num)){
        return formula.error.v;
      }

      start_num = parseInt(start_num);

      //要在文本中替换的字符个数
      let num_chars = func_methods.getFirstValue(arguments[2]);
      if(valueIsError(num_chars)){
        return num_chars;
      }

      if(!isRealNum(num_chars)){
        return formula.error.v;
      }

      num_chars = parseInt(num_chars);

      //字符串2
      let new_text = func_methods.getFirstValue(arguments[3], 'text');
      if(valueIsError(new_text)){
        return new_text;
      }

      new_text = new_text.toString();

      return old_text.substr(0, start_num - 1) + new_text + old_text.substr(start_num - 1 + num_chars);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REPT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串1
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      //重复次数
      let number_times = func_methods.getFirstValue(arguments[1]);
      if(valueIsError(number_times)){
        return number_times;
      }

      if(!isRealNum(number_times)){
        return formula.error.v;
      }

      number_times = parseInt(number_times);

      if(number_times < 0){
        return formula.error.v;
      }

      if(number_times > 100){
        number_times = 100;
      }

      return new Array(number_times + 1).join(text);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SEARCH': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串1
      let find_text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(find_text)){
        return find_text;
      }

      find_text = find_text.toString();

      //字符串2
      let within_text = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(within_text)){
        return within_text;
      }

      within_text = within_text.toString();

      //开始位置
      let start_num = 1;
      if(arguments.length == 3){
        start_num = func_methods.getFirstValue(arguments[2]);
        if(valueIsError(start_num)){
          return start_num;
        }

        if(!isRealNum(start_num)){
          return formula.error.v;
        }

        start_num = parseInt(start_num);
      }

      if(start_num <= 0 || start_num > within_text.length){
        return formula.error.v;
      }

      const foundAt = within_text.toLowerCase().indexOf(find_text.toLowerCase(), start_num - 1) + 1;

      return (foundAt === 0) ? formula.error.v : foundAt;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SUBSTITUTE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //需要替换其中字符的文本
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      //需要替换的文本
      let old_text = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(old_text)){
        return old_text;
      }

      old_text = old_text.toString();

      //用于替换 old_text 的文本
      let new_text = func_methods.getFirstValue(arguments[2], 'text');
      if(valueIsError(new_text)){
        return new_text;
      }

      new_text = new_text.toString();

      //instance_num
      let instance_num = null;
      if(arguments.length == 4){
        instance_num = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(instance_num)){
          return instance_num;
        }

        if(!isRealNum(instance_num)){
          return formula.error.v;
        }

        instance_num = parseInt(instance_num);
      }

      //计算
      const reg = new RegExp(old_text, 'g');

      let result;

      if(instance_num == null){
        result = text.replace(reg, new_text);
      }
      else{
        if(instance_num <= 0){
          return formula.error.v;
        }

        const match = text.match(reg);

        if(match == null || instance_num > match.length){
          return text;
        }
        else{
          const len = old_text.length;
          let index = 0;

          for(var i = 1; i <= instance_num; i++){
            index = text.indexOf(old_text, index) + 1;
          }

          result = text.substring(0, index - 1) + new_text + text.substring(index - 1 + len);
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CLEAN': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      const textArr = [];
      for(var i = 0; i < text.length; i++){
        const code = text.charCodeAt(i);

        if(/[\u4e00-\u9fa5]/g.test(text.charAt(i)) || (code > 31 && code < 127)){
          textArr.push(text.charAt(i));
        }
      }

      return textArr.join('');
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TEXT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let value = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value)){
        return value;
      }

      if(!isRealNum(value)){
        return formula.error.v;
      }

      value = parseFloat(value);

      //格式
      let format_text = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(format_text)){
        return format_text;
      }

      format_text = format_text.toString();

      return update(format_text, value);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TRIM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      return text.replace(/ +/g, ' ').trim();
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'VALUE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString();

      return genarate(text)[2];
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PROPER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //字符串
      let text = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(text)){
        return text;
      }

      text = text.toString().toLowerCase();

      return text.replace(/[a-zA-Z]+/g, function(word){ return word.substring(0,1).toUpperCase() + word.substring(1); });
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CONVERT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let number = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(number)){
        return number;
      }

      if(!isRealNum(number)){
        return formula.error.v;
      }

      number = parseFloat(number);

      //数值的单位
      let from_unit = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(from_unit)){
        return from_unit;
      }

      from_unit = from_unit.toString();

      //结果的单位
      let to_unit = func_methods.getFirstValue(arguments[2], 'text');
      if(valueIsError(to_unit)){
        return to_unit;
      }

      to_unit = to_unit.toString();

      //计算
      const units = [
        ['a.u. of action', '?', null, 'action', false, false, 1.05457168181818e-34],
        ['a.u. of charge', 'e', null, 'electric_charge', false, false, 1.60217653141414e-19],
        ['a.u. of energy', 'Eh', null, 'energy', false, false, 4.35974417757576e-18],
        ['a.u. of length', 'a?', null, 'length', false, false, 5.29177210818182e-11],
        ['a.u. of mass', 'm?', null, 'mass', false, false, 9.10938261616162e-31],
        ['a.u. of time', '?/Eh', null, 'time', false, false, 2.41888432650516e-17],
        ['admiralty knot', 'admkn', null, 'speed', false, true, 0.514773333],
        ['ampere', 'A', null, 'electric_current', true, false, 1],
        ['ampere per meter', 'A/m', null, 'magnetic_field_intensity', true, false, 1],
        ['ångström', 'Å', ['ang'], 'length', false, true, 1e-10],
        ['are', 'ar', null, 'area', false, true, 100],
        ['astronomical unit', 'ua', null, 'length', false, false, 1.49597870691667e-11],
        ['bar', 'bar', null, 'pressure', false, false, 100000],
        ['barn', 'b', null, 'area', false, false, 1e-28],
        ['becquerel', 'Bq', null, 'radioactivity', true, false, 1],
        ['bit', 'bit', ['b'], 'information', false, true, 1],
        ['btu', 'BTU', ['btu'], 'energy', false, true, 1055.05585262],
        ['byte', 'byte', null, 'information', false, true, 8],
        ['candela', 'cd', null, 'luminous_intensity', true, false, 1],
        ['candela per square metre', 'cd/m?', null, 'luminance', true, false, 1],
        ['coulomb', 'C', null, 'electric_charge', true, false, 1],
        ['cubic ångström', 'ang3', ['ang^3'], 'volume', false, true, 1e-30],
        ['cubic foot', 'ft3', ['ft^3'], 'volume', false, true, 0.028316846592],
        ['cubic inch', 'in3', ['in^3'], 'volume', false, true, 0.000016387064],
        ['cubic light-year', 'ly3', ['ly^3'], 'volume', false, true, 8.46786664623715e-47],
        ['cubic metre', 'm?', null, 'volume', true, true, 1],
        ['cubic mile', 'mi3', ['mi^3'], 'volume', false, true, 4168181825.44058],
        ['cubic nautical mile', 'Nmi3', ['Nmi^3'], 'volume', false, true, 6352182208],
        ['cubic Pica', 'Pica3', ['Picapt3', 'Pica^3', 'Picapt^3'], 'volume', false, true, 7.58660370370369e-8],
        ['cubic yard', 'yd3', ['yd^3'], 'volume', false, true, 0.764554857984],
        ['cup', 'cup', null, 'volume', false, true, 0.0002365882365],
        ['dalton', 'Da', ['u'], 'mass', false, false, 1.66053886282828e-27],
        ['day', 'd', ['day'], 'time', false, true, 86400],
        ['degree', '°', null, 'angle', false, false, 0.0174532925199433],
        ['degrees Rankine', 'Rank', null, 'temperature', false, true, 0.555555555555556],
        ['dyne', 'dyn', ['dy'], 'force', false, true, 0.00001],
        ['electronvolt', 'eV', ['ev'], 'energy', false, true, 1.60217656514141],
        ['ell', 'ell', null, 'length', false, true, 1.143],
        ['erg', 'erg', ['e'], 'energy', false, true, 1e-7],
        ['farad', 'F', null, 'electric_capacitance', true, false, 1],
        ['fluid ounce', 'oz', null, 'volume', false, true, 0.0000295735295625],
        ['foot', 'ft', null, 'length', false, true, 0.3048],
        ['foot-pound', 'flb', null, 'energy', false, true, 1.3558179483314],
        ['gal', 'Gal', null, 'acceleration', false, false, 0.01],
        ['gallon', 'gal', null, 'volume', false, true, 0.003785411784],
        ['gauss', 'G', ['ga'], 'magnetic_flux_density', false, true, 1],
        ['grain', 'grain', null, 'mass', false, true, 0.0000647989],
        ['gram', 'g', null, 'mass', false, true, 0.001],
        ['gray', 'Gy', null, 'absorbed_dose', true, false, 1],
        ['gross registered ton', 'GRT', ['regton'], 'volume', false, true, 2.8316846592],
        ['hectare', 'ha', null, 'area', false, true, 10000],
        ['henry', 'H', null, 'inductance', true, false, 1],
        ['hertz', 'Hz', null, 'frequency', true, false, 1],
        ['horsepower', 'HP', ['h'], 'power', false, true, 745.69987158227],
        ['horsepower-hour', 'HPh', ['hh', 'hph'], 'energy', false, true, 2684519.538],
        ['hour', 'h', ['hr'], 'time', false, true, 3600],
        ['imperial gallon (U.K.)', 'uk_gal', null, 'volume', false, true, 0.00454609],
        ['imperial hundredweight', 'lcwt', ['uk_cwt', 'hweight'], 'mass', false, true, 50.802345],
        ['imperial quart (U.K)', 'uk_qt', null, 'volume', false, true, 0.0011365225],
        ['imperial ton', 'brton', ['uk_ton', 'LTON'], 'mass', false, true, 1016.046909],
        ['inch', 'in', null, 'length', false, true, 0.0254],
        ['international acre', 'uk_acre', null, 'area', false, true, 4046.8564224],
        ['IT calorie', 'cal', null, 'energy', false, true, 4.1868],
        ['joule', 'J', null, 'energy', true, true, 1],
        ['katal', 'kat', null, 'catalytic_activity', true, false, 1],
        ['kelvin', 'K', ['kel'], 'temperature', true, true, 1],
        ['kilogram', 'kg', null, 'mass', true, true, 1],
        ['knot', 'kn', null, 'speed', false, true, 0.514444444444444],
        ['light-year', 'ly', null, 'length', false, true, 9460730472580800],
        ['litre', 'L', ['l', 'lt'], 'volume', false, true, 0.001],
        ['lumen', 'lm', null, 'luminous_flux', true, false, 1],
        ['lux', 'lx', null, 'illuminance', true, false, 1],
        ['maxwell', 'Mx', null, 'magnetic_flux', false, false, 1e-18],
        ['measurement ton', 'MTON', null, 'volume', false, true, 1.13267386368],
        ['meter per hour', 'm/h', ['m/hr'], 'speed', false, true, 0.00027777777777778],
        ['meter per second', 'm/s', ['m/sec'], 'speed', true, true, 1],
        ['meter per second squared', 'm?s??', null, 'acceleration', true, false, 1],
        ['parsec', 'pc', ['parsec'], 'length', false, true, 30856775814671900],
        ['meter squared per second', 'm?/s', null, 'kinematic_viscosity', true, false, 1],
        ['metre', 'm', null, 'length', true, true, 1],
        ['miles per hour', 'mph', null, 'speed', false, true, 0.44704],
        ['millimetre of mercury', 'mmHg', null, 'pressure', false, false, 133.322],
        ['minute', '?', null, 'angle', false, false, 0.000290888208665722],
        ['minute', 'min', ['mn'], 'time', false, true, 60],
        ['modern teaspoon', 'tspm', null, 'volume', false, true, 0.000005],
        ['mole', 'mol', null, 'amount_of_substance', true, false, 1],
        ['morgen', 'Morgen', null, 'area', false, true, 2500],
        ['n.u. of action', '?', null, 'action', false, false, 1.05457168181818e-34],
        ['n.u. of mass', 'm?', null, 'mass', false, false, 9.10938261616162e-31],
        ['n.u. of speed', 'c?', null, 'speed', false, false, 299792458],
        ['n.u. of time', '?/(me?c??)', null, 'time', false, false, 1.28808866778687e-21],
        ['nautical mile', 'M', ['Nmi'], 'length', false, true, 1852],
        ['newton', 'N', null, 'force', true, true, 1],
        ['œrsted', 'Oe ', null, 'magnetic_field_intensity', false, false, 79.5774715459477],
        ['ohm', 'Ω', null, 'electric_resistance', true, false, 1],
        ['ounce mass', 'ozm', null, 'mass', false, true, 0.028349523125],
        ['pascal', 'Pa', null, 'pressure', true, false, 1],
        ['pascal second', 'Pa?s', null, 'dynamic_viscosity', true, false, 1],
        ['pferdestärke', 'PS', null, 'power', false, true, 735.49875],
        ['phot', 'ph', null, 'illuminance', false, false, 0.0001],
        ['pica (1/6 inch)', 'pica', null, 'length', false, true, 0.00035277777777778],
        ['pica (1/72 inch)', 'Pica', ['Picapt'], 'length', false, true, 0.00423333333333333],
        ['poise', 'P', null, 'dynamic_viscosity', false, false, 0.1],
        ['pond', 'pond', null, 'force', false, true, 0.00980665],
        ['pound force', 'lbf', null, 'force', false, true, 4.4482216152605],
        ['pound mass', 'lbm', null, 'mass', false, true, 0.45359237],
        ['quart', 'qt', null, 'volume', false, true, 0.000946352946],
        ['radian', 'rad', null, 'angle', true, false, 1],
        ['second', '?', null, 'angle', false, false, 0.00000484813681109536],
        ['second', 's', ['sec'], 'time', true, true, 1],
        ['short hundredweight', 'cwt', ['shweight'], 'mass', false, true, 45.359237],
        ['siemens', 'S', null, 'electrical_conductance', true, false, 1],
        ['sievert', 'Sv', null, 'equivalent_dose', true, false, 1],
        ['slug', 'sg', null, 'mass', false, true, 14.59390294],
        ['square ångström', 'ang2', ['ang^2'], 'area', false, true, 1e-20],
        ['square foot', 'ft2', ['ft^2'], 'area', false, true, 0.09290304],
        ['square inch', 'in2', ['in^2'], 'area', false, true, 0.00064516],
        ['square light-year', 'ly2', ['ly^2'], 'area', false, true, 8.95054210748189e+31],
        ['square meter', 'm?', null, 'area', true, true, 1],
        ['square mile', 'mi2', ['mi^2'], 'area', false, true, 2589988.110336],
        ['square nautical mile', 'Nmi2', ['Nmi^2'], 'area', false, true, 3429904],
        ['square Pica', 'Pica2', ['Picapt2', 'Pica^2', 'Picapt^2'], 'area', false, true, 0.00001792111111111],
        ['square yard', 'yd2', ['yd^2'], 'area', false, true, 0.83612736],
        ['statute mile', 'mi', null, 'length', false, true, 1609.344],
        ['steradian', 'sr', null, 'solid_angle', true, false, 1],
        ['stilb', 'sb', null, 'luminance', false, false, 0.0001],
        ['stokes', 'St', null, 'kinematic_viscosity', false, false, 0.0001],
        ['stone', 'stone', null, 'mass', false, true, 6.35029318],
        ['tablespoon', 'tbs', null, 'volume', false, true, 0.0000147868],
        ['teaspoon', 'tsp', null, 'volume', false, true, 0.00000492892],
        ['tesla', 'T', null, 'magnetic_flux_density', true, true, 1],
        ['thermodynamic calorie', 'c', null, 'energy', false, true, 4.184],
        ['ton', 'ton', null, 'mass', false, true, 907.18474],
        ['tonne', 't', null, 'mass', false, false, 1000],
        ['U.K. pint', 'uk_pt', null, 'volume', false, true, 0.00056826125],
        ['U.S. bushel', 'bushel', null, 'volume', false, true, 0.03523907],
        ['U.S. oil barrel', 'barrel', null, 'volume', false, true, 0.158987295],
        ['U.S. pint', 'pt', ['us_pt'], 'volume', false, true, 0.000473176473],
        ['U.S. survey mile', 'survey_mi', null, 'length', false, true, 1609.347219],
        ['U.S. survey/statute acre', 'us_acre', null, 'area', false, true, 4046.87261],
        ['volt', 'V', null, 'voltage', true, false, 1],
        ['watt', 'W', null, 'power', true, true, 1],
        ['watt-hour', 'Wh', ['wh'], 'energy', false, true, 3600],
        ['weber', 'Wb', null, 'magnetic_flux', true, false, 1],
        ['yard', 'yd', null, 'length', false, true, 0.9144],
        ['year', 'yr', null, 'time', false, true, 31557600]
      ];

      const binary_prefixes = {
        Yi: ['yobi', 80, 1208925819614629174706176, 'Yi', 'yotta'],
        Zi: ['zebi', 70, 1180591620717411303424, 'Zi', 'zetta'],
        Ei: ['exbi', 60, 1152921504606846976, 'Ei', 'exa'],
        Pi: ['pebi', 50, 1125899906842624, 'Pi', 'peta'],
        Ti: ['tebi', 40, 1099511627776, 'Ti', 'tera'],
        Gi: ['gibi', 30, 1073741824, 'Gi', 'giga'],
        Mi: ['mebi', 20, 1048576, 'Mi', 'mega'],
        ki: ['kibi', 10, 1024, 'ki', 'kilo']
      };

      const unit_prefixes = {
        Y: ['yotta', 1e+24, 'Y'],
        Z: ['zetta', 1e+21, 'Z'],
        E: ['exa', 1e+18, 'E'],
        P: ['peta', 1e+15, 'P'],
        T: ['tera', 1e+12, 'T'],
        G: ['giga', 1e+09, 'G'],
        M: ['mega', 1e+06, 'M'],
        k: ['kilo', 1e+03, 'k'],
        h: ['hecto', 1e+02, 'h'],
        e: ['dekao', 1e+01, 'e'],
        d: ['deci', 1e-01, 'd'],
        c: ['centi', 1e-02, 'c'],
        m: ['milli', 1e-03, 'm'],
        u: ['micro', 1e-06, 'u'],
        n: ['nano', 1e-09, 'n'],
        p: ['pico', 1e-12, 'p'],
        f: ['femto', 1e-15, 'f'],
        a: ['atto', 1e-18, 'a'],
        z: ['zepto', 1e-21, 'z'],
        y: ['yocto', 1e-24, 'y']
      };

      let from = null;
      let to = null;
      let base_from_unit = from_unit;
      let base_to_unit = to_unit;
      let from_multiplier = 1;
      let to_multiplier = 1;
      let alt;

      for (var i = 0; i < units.length; i++) {
        alt = (units[i][2] === null) ? [] : units[i][2];

        if (units[i][1] === base_from_unit || alt.indexOf(base_from_unit) >= 0) {
          from = units[i];
        }

        if (units[i][1] === base_to_unit || alt.indexOf(base_to_unit) >= 0) {
          to = units[i];
        }
      }

      if (from === null) {
        const from_binary_prefix = binary_prefixes[from_unit.substring(0, 2)];
        let from_unit_prefix = unit_prefixes[from_unit.substring(0, 1)];

        if (from_unit.substring(0, 2) === 'da') {
          from_unit_prefix = ['dekao', 1e+01, 'da'];
        }

        if (from_binary_prefix) {
          from_multiplier = from_binary_prefix[2];
          base_from_unit = from_unit.substring(2);
        }
        else if (from_unit_prefix) {
          from_multiplier = from_unit_prefix[1];
          base_from_unit = from_unit.substring(from_unit_prefix[2].length);
        }

        for (let j = 0; j < units.length; j++) {
          alt = (units[j][2] === null) ? [] : units[j][2];

          if (units[j][1] === base_from_unit || alt.indexOf(base_from_unit) >= 0) {
            from = units[j];
          }
        }
      }

      if (to === null) {
        const to_binary_prefix = binary_prefixes[to_unit.substring(0, 2)];
        let to_unit_prefix = unit_prefixes[to_unit.substring(0, 1)];

        if (to_unit.substring(0, 2) === 'da') {
          to_unit_prefix = ['dekao', 1e+01, 'da'];
        }

        if (to_binary_prefix) {
          to_multiplier = to_binary_prefix[2];
          base_to_unit = to_unit.substring(2);
        }
        else if (to_unit_prefix) {
          to_multiplier = to_unit_prefix[1];
          base_to_unit = to_unit.substring(to_unit_prefix[2].length);
        }

        for (let k = 0; k < units.length; k++) {
          alt = (units[k][2] === null) ? [] : units[k][2];

          if (units[k][1] === base_to_unit || alt.indexOf(base_to_unit) >= 0) {
            to = units[k];
          }
        }
      }

      if (from === null || to === null) {
        return formula.error.na;
      }

      if (from[3] !== to[3]) {
        return formula.error.na;
      }

      return number * from[6] * from_multiplier / (to[6] * to_multiplier);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SUMX2MY2': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //第一个数组或数值区域
      const data_array_x = arguments[0];
      let array_x = [];

      if(getObjType(data_array_x) == 'array'){
        if(getObjType(data_array_x[0]) == 'array' && !func_methods.isDyadicArr(data_array_x)){
          return formula.error.v;
        }

        array_x = array_x.concat(func_methods.getDataArr(data_array_x, false));
      }
      else if(getObjType(data_array_x) == 'object' && data_array_x.startCell != null){
        array_x = array_x.concat(func_methods.getCellDataArr(data_array_x, 'text', false));
      }
      else{
        array_x.push(data_array_x);
      }

      //第二个数组或数值区域
      const data_array_y = arguments[1];
      let array_y = [];

      if(getObjType(data_array_y) == 'array'){
        if(getObjType(data_array_y[0]) == 'array' && !func_methods.isDyadicArr(data_array_y)){
          return formula.error.v;
        }

        array_y = array_y.concat(func_methods.getDataArr(data_array_y, false));
      }
      else if(getObjType(data_array_y) == 'object' && data_array_y.startCell != null){
        array_y = array_y.concat(func_methods.getCellDataArr(data_array_y, 'text', false));
      }
      else{
        array_y.push(data_array_y);
      }

      if(array_x.length != array_y.length){
        return formula.error.na;
      }

      //array_x 和 array_y 只取数值
      const data_x = [], data_y = [];

      for(var i = 0; i < array_x.length; i++){
        const num_x = array_x[i];
        const num_y = array_y[i];

        if(isRealNum(num_x) && isRealNum(num_y)){
          data_x.push(parseFloat(num_x));
          data_y.push(parseFloat(num_y));
        }
      }

      //计算
      let sum = 0;

      for (var i = 0; i < data_x.length; i++) {
        sum += Math.pow(data_x[i], 2) - Math.pow(data_y[i], 2);
      }

      return sum;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SUMX2PY2': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //第一个数组或数值区域
      const data_array_x = arguments[0];
      let array_x = [];

      if(getObjType(data_array_x) == 'array'){
        if(getObjType(data_array_x[0]) == 'array' && !func_methods.isDyadicArr(data_array_x)){
          return formula.error.v;
        }

        array_x = array_x.concat(func_methods.getDataArr(data_array_x, false));
      }
      else if(getObjType(data_array_x) == 'object' && data_array_x.startCell != null){
        array_x = array_x.concat(func_methods.getCellDataArr(data_array_x, 'text', false));
      }
      else{
        array_x.push(data_array_x);
      }

      //第二个数组或数值区域
      const data_array_y = arguments[1];
      let array_y = [];

      if(getObjType(data_array_y) == 'array'){
        if(getObjType(data_array_y[0]) == 'array' && !func_methods.isDyadicArr(data_array_y)){
          return formula.error.v;
        }

        array_y = array_y.concat(func_methods.getDataArr(data_array_y, false));
      }
      else if(getObjType(data_array_y) == 'object' && data_array_y.startCell != null){
        array_y = array_y.concat(func_methods.getCellDataArr(data_array_y, 'text', false));
      }
      else{
        array_y.push(data_array_y);
      }

      if(array_x.length != array_y.length){
        return formula.error.na;
      }

      //array_x 和 array_y 只取数值
      const data_x = [], data_y = [];

      for(var i = 0; i < array_x.length; i++){
        const num_x = array_x[i];
        const num_y = array_y[i];

        if(isRealNum(num_x) && isRealNum(num_y)){
          data_x.push(parseFloat(num_x));
          data_y.push(parseFloat(num_y));
        }
      }

      //计算
      let sum = 0;

      for (var i = 0; i < data_x.length; i++) {
        sum += Math.pow(data_x[i], 2) + Math.pow(data_y[i], 2);
      }

      return sum;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SUMXMY2': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //第一个数组或数值区域
      const data_array_x = arguments[0];
      let array_x = [];

      if(getObjType(data_array_x) == 'array'){
        if(getObjType(data_array_x[0]) == 'array' && !func_methods.isDyadicArr(data_array_x)){
          return formula.error.v;
        }

        array_x = array_x.concat(func_methods.getDataArr(data_array_x, false));
      }
      else if(getObjType(data_array_x) == 'object' && data_array_x.startCell != null){
        array_x = array_x.concat(func_methods.getCellDataArr(data_array_x, 'text', false));
      }
      else{
        array_x.push(data_array_x);
      }

      //第二个数组或数值区域
      const data_array_y = arguments[1];
      let array_y = [];

      if(getObjType(data_array_y) == 'array'){
        if(getObjType(data_array_y[0]) == 'array' && !func_methods.isDyadicArr(data_array_y)){
          return formula.error.v;
        }

        array_y = array_y.concat(func_methods.getDataArr(data_array_y, false));
      }
      else if(getObjType(data_array_y) == 'object' && data_array_y.startCell != null){
        array_y = array_y.concat(func_methods.getCellDataArr(data_array_y, 'text', false));
      }
      else{
        array_y.push(data_array_y);
      }

      if(array_x.length != array_y.length){
        return formula.error.na;
      }

      //array_x 和 array_y 只取数值
      const data_x = [], data_y = [];

      for(var i = 0; i < array_x.length; i++){
        const num_x = array_x[i];
        const num_y = array_y[i];

        if(isRealNum(num_x) && isRealNum(num_y)){
          data_x.push(parseFloat(num_x));
          data_y.push(parseFloat(num_y));
        }
      }

      //计算
      let sum = 0;

      for (var i = 0; i < data_x.length; i++) {
        sum += Math.pow(data_x[i] - data_y[i], 2);
      }

      return sum;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TRANSPOSE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //从其返回唯一值的数组或区域
      const data_array = arguments[0];
      let array = [];

      if(getObjType(data_array) == 'array'){
        if(getObjType(data_array[0]) == 'array' && !func_methods.isDyadicArr(data_array)){
          return formula.error.v;
        }

        array = func_methods.getDataDyadicArr(data_array);
      }
      else if(getObjType(data_array) == 'object' && data_array.startCell != null){
        array = func_methods.getCellDataDyadicArr(data_array, 'number');
      }

      array = array[0].map(function(col, a){
        return array.map(function(row){
          return row[a];
        });
      });

      return array;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TREND': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //已知的 y 值集合
      const data_known_y = arguments[0];
      let known_y = [];

      if(getObjType(data_known_y) == 'array'){
        if(getObjType(data_known_y[0]) == 'array' && !func_methods.isDyadicArr(data_known_y)){
          return formula.error.v;
        }

        known_y = func_methods.getDataDyadicArr(data_known_y);
      }
      else if(getObjType(data_known_y) == 'object' && data_known_y.startCell != null){
        known_y = func_methods.getCellDataDyadicArr(data_known_y, 'text');
      }
      else{
        if(!isRealNum(data_known_y)){
          return formula.error.v;
        }

        var rowArr = [];

        rowArr.push(parseFloat(data_known_y));

        known_y.push(rowArr);
      }

      const known_y_rowlen = known_y.length;
      const known_y_collen = known_y[0].length;

      for(var i = 0; i < known_y_rowlen; i++){
        for(var j = 0; j < known_y_collen; j++){
          if(!isRealNum(known_y[i][j])){
            return formula.error.v;
          }

          known_y[i][j] = parseFloat(known_y[i][j]);
        }
      }

      //可选 x 值集合
      let known_x = [];
      for(var i = 1; i <= known_y_rowlen; i++){
        for(var j = 1; j <= known_y_collen; j++){
          const number = (i - 1) * known_y_collen + j;
          known_x.push(number);
        }
      }

      if(arguments.length >= 2){
        const data_known_x = arguments[1];
        known_x = [];

        if(getObjType(data_known_x) == 'array'){
          if(getObjType(data_known_x[0]) == 'array' && !func_methods.isDyadicArr(data_known_x)){
            return formula.error.v;
          }

          known_x = func_methods.getDataDyadicArr(data_known_x);
        }
        else if(getObjType(data_known_x) == 'object' && data_known_x.startCell != null){
          known_x = func_methods.getCellDataDyadicArr(data_known_x, 'text');
        }
        else{
          if(!isRealNum(data_known_x)){
            return formula.error.v;
          }

          var rowArr = [];

          rowArr.push(parseFloat(data_known_x));

          known_x.push(rowArr);
        }

        for(var i = 0; i < known_x.length; i++){
          for(var j = 0; j < known_x[0].length; j++){
            if(!isRealNum(known_x[i][j])){
              return formula.error.v;
            }

            known_x[i][j] = parseFloat(known_x[i][j]);
          }
        }
      }

      const known_x_rowlen = known_x.length;
      const known_x_collen = known_x[0].length;

      //新 x 值
      let new_x = known_x;

      if(arguments.length >= 3){
        const data_new_x = arguments[2];
        new_x = [];

        if(getObjType(data_new_x) == 'array'){
          if(getObjType(data_new_x[0]) == 'array' && !func_methods.isDyadicArr(data_new_x)){
            return formula.error.v;
          }

          new_x = func_methods.getDataDyadicArr(data_new_x);
        }
        else if(getObjType(data_new_x) == 'object' && data_new_x.startCell != null){
          new_x = func_methods.getCellDataDyadicArr(data_new_x, 'text');
        }
        else{
          if(!isRealNum(data_new_x)){
            return formula.error.v;
          }

          var rowArr = [];

          rowArr.push(parseFloat(data_new_x));

          new_x.push(rowArr);
        }

        for(var i = 0; i < new_x.length; i++){
          for(var j = 0; j < new_x[0].length; j++){
            if(!isRealNum(new_x[i][j])){
              return formula.error.v;
            }

            new_x[i][j] = parseFloat(new_x[i][j]);
          }
        }
      }

      //逻辑值
      let const_b = true;

      if(arguments.length == 4){
        const_b = func_methods.getCellBoolen(arguments[3]);

        if(valueIsError(const_b)){
          return const_b;
        }
      }

      if(known_y_rowlen != known_x_rowlen || known_y_collen != known_x_collen){
        return formula.error.r;
      }

      //计算
      function leastSquare(arr_x, arr_y){
        let xSum = 0, ySum = 0, xySum = 0, x2Sum = 0;

        for(let i = 0; i < arr_x.length; i++){
          for(let j = 0; j < arr_x[i].length; j++){
            xSum += arr_x[i][j];
            ySum += arr_y[i][j];
            xySum += arr_x[i][j] * arr_y[i][j];
            x2Sum += arr_x[i][j] * arr_x[i][j];
          }
        }

        const n = arr_x.length * arr_x[0].length;

        const xMean = xSum / n;
        const yMean = ySum / n;
        const xyMean = xySum / n;
        const x2Mean = x2Sum / n;

        const m = (xyMean - xMean * yMean) / (x2Mean - xMean * xMean);
        const b = yMean - m * xMean;

        return [m, b];
      }

      const ls = leastSquare(known_x, known_y);
      const m = ls[0];

      if(const_b){
        var b = ls[1];
      }
      else{
        var b = 0;
      }

      const result = [];

      for(var i = 0; i < new_x.length; i++){
        for(var j = 0; j < new_x[i].length; j++){
          const x = new_x[i][j];
          const y = m * x + b;

          result.push(Math.round(y * 1000000000) / 1000000000);
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'FREQUENCY': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //频率数组
      const data_data_array = arguments[0];
      let data_array = [];

      if(getObjType(data_data_array) == 'array'){
        if(getObjType(data_data_array[0]) == 'array' && !func_methods.isDyadicArr(data_data_array)){
          return formula.error.v;
        }

        data_array = data_array.concat(func_methods.getDataArr(data_data_array, true));
      }
      else if(getObjType(data_data_array) == 'object' && data_data_array.startCell != null){
        data_array = data_array.concat(func_methods.getCellDataArr(data_data_array, 'number', true));
      }
      else{
        if(!isRealNum(data_data_array)){
          return formula.error.v;
        }

        data_array.push(data_data_array);
      }

      const data_array_n = [];

      for(var i = 0; i < data_array.length; i++){
        if(isRealNum(data_array[i])){
          data_array_n.push(parseFloat(data_array[i]));
        }
      }

      //间隔数组
      const data_bins_array = arguments[1];
      let bins_array = [];

      if(getObjType(data_bins_array) == 'array'){
        if(getObjType(data_bins_array[0]) == 'array' && !func_methods.isDyadicArr(data_bins_array)){
          return formula.error.v;
        }

        bins_array = bins_array.concat(func_methods.getDataArr(data_bins_array, true));
      }
      else if(getObjType(data_bins_array) == 'object' && data_bins_array.startCell != null){
        bins_array = bins_array.concat(func_methods.getCellDataArr(data_bins_array, 'number', true));
      }
      else{
        if(!isRealNum(data_bins_array)){
          return formula.error.v;
        }

        bins_array.push(data_bins_array);
      }

      const bins_array_n = [];

      for(var i = 0; i < bins_array.length; i++){
        if(isRealNum(bins_array[i])){
          bins_array_n.push(parseFloat(bins_array[i]));
        }
      }

      //计算
      if(data_array_n.length == 0 && bins_array_n.length == 0){
        return [[0], [0]];
      }
      else if(data_array_n.length == 0){
        var result = [[0]];

        for(var i = 0; i < bins_array_n.length; i++){
          result.push([0]);
        }

        return result;
      }
      else if(bins_array_n.length == 0){
        return [[0], [data_array_n.length]];
      }
      else{
        bins_array_n.sort(function(a, b){
          return a - b;
        });

        var result = [];

        for(var i = 0; i < bins_array_n.length; i++){
          if(i == 0){
            var count = 0;

            for(var j = 0; j < data_array_n.length; j++){
              if(data_array_n[j] <= bins_array_n[0]){
                count++;
              }
            }

            result.push([count]);
          }
          else if(i == bins_array_n.length - 1){
            let count1 = 0, count2 = 0;

            for(var j = 0; j < data_array_n.length; j++){
              if(data_array_n[j] <= bins_array_n[i] && data_array_n[j] > bins_array_n[i - 1]){
                count1++;
              }

              if(data_array_n[j] > bins_array_n[i]){
                count2++;
              }
            }

            result.push([count1]);
            result.push([count2]);
          }
          else{
            var count = 0;

            for(var j = 0; j < data_array_n.length; j++){
              if(data_array_n[j] <= bins_array_n[i] && data_array_n[j] > bins_array_n[i - 1]){
                count++;
              }
            }

            result.push([count]);
          }
        }

        return result;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'GROWTH': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //已知的 y 值集合
      const data_known_y = arguments[0];
      let known_y = [];

      if(getObjType(data_known_y) == 'array'){
        if(getObjType(data_known_y[0]) == 'array' && !func_methods.isDyadicArr(data_known_y)){
          return formula.error.v;
        }

        known_y = func_methods.getDataDyadicArr(data_known_y);
      }
      else if(getObjType(data_known_y) == 'object' && data_known_y.startCell != null){
        known_y = func_methods.getCellDataDyadicArr(data_known_y, 'text');
      }
      else{
        if(!isRealNum(data_known_y)){
          return formula.error.v;
        }

        var rowArr = [];

        rowArr.push(parseFloat(data_known_y));

        known_y.push(rowArr);
      }

      const known_y_rowlen = known_y.length;
      const known_y_collen = known_y[0].length;

      for(var i = 0; i < known_y_rowlen; i++){
        for(var j = 0; j < known_y_collen; j++){
          if(!isRealNum(known_y[i][j])){
            return formula.error.v;
          }

          known_y[i][j] = parseFloat(known_y[i][j]);
        }
      }

      //可选 x 值集合
      let known_x = [];
      for(var i = 1; i <= known_y_rowlen; i++){
        for(var j = 1; j <= known_y_collen; j++){
          const number = (i - 1) * known_y_collen + j;
          known_x.push(number);
        }
      }

      if(arguments.length >= 2){
        const data_known_x = arguments[1];
        known_x = [];

        if(getObjType(data_known_x) == 'array'){
          if(getObjType(data_known_x[0]) == 'array' && !func_methods.isDyadicArr(data_known_x)){
            return formula.error.v;
          }

          known_x = func_methods.getDataDyadicArr(data_known_x);
        }
        else if(getObjType(data_known_x) == 'object' && data_known_x.startCell != null){
          known_x = func_methods.getCellDataDyadicArr(data_known_x, 'text');
        }
        else{
          if(!isRealNum(data_known_x)){
            return formula.error.v;
          }

          var rowArr = [];

          rowArr.push(parseFloat(data_known_x));

          known_x.push(rowArr);
        }

        for(var i = 0; i < known_x.length; i++){
          for(var j = 0; j < known_x[0].length; j++){
            if(!isRealNum(known_x[i][j])){
              return formula.error.v;
            }

            known_x[i][j] = parseFloat(known_x[i][j]);
          }
        }
      }

      const known_x_rowlen = known_x.length;
      const known_x_collen = known_x[0].length;

      //新 x 值
      let new_x = known_x;

      if(arguments.length >= 3){
        const data_new_x = arguments[2];
        new_x = [];

        if(getObjType(data_new_x) == 'array'){
          if(getObjType(data_new_x[0]) == 'array' && !func_methods.isDyadicArr(data_new_x)){
            return formula.error.v;
          }

          new_x = func_methods.getDataDyadicArr(data_new_x);
        }
        else if(getObjType(data_new_x) == 'object' && data_new_x.startCell != null){
          new_x = func_methods.getCellDataDyadicArr(data_new_x, 'text');
        }
        else{
          if(!isRealNum(data_new_x)){
            return formula.error.v;
          }

          var rowArr = [];

          rowArr.push(parseFloat(data_new_x));

          new_x.push(rowArr);
        }

        for(var i = 0; i < new_x.length; i++){
          for(var j = 0; j < new_x[0].length; j++){
            if(!isRealNum(new_x[i][j])){
              return formula.error.v;
            }

            new_x[i][j] = parseFloat(new_x[i][j]);
          }
        }
      }

      //逻辑值
      let const_b = true;

      if(arguments.length == 4){
        const_b = func_methods.getCellBoolen(arguments[3]);

        if(valueIsError(const_b)){
          return const_b;
        }
      }

      if(known_y_rowlen != known_x_rowlen || known_y_collen != known_x_collen){
        return formula.error.r;
      }

      //计算
      function leastSquare(arr_x, arr_y){
        let xSum = 0, ySum = 0, xySum = 0, x2Sum = 0;

        for(let i = 0; i < arr_x.length; i++){
          for(let j = 0; j < arr_x[i].length; j++){
            xSum += arr_x[i][j];
            // ySum += arr_y[i][j];
            ySum += Math.log(arr_y[i][j]);
            // xySum += arr_x[i][j] * arr_y[i][j];
            xySum += arr_x[i][j] * Math.log(arr_y[i][j]);
            x2Sum += arr_x[i][j] * arr_x[i][j];
          }
        }

        const n = arr_x.length * arr_x[0].length;

        const xMean = xSum / n;
        const yMean = ySum / n;
        const xyMean = xySum / n;
        const x2Mean = x2Sum / n;

        const m = (xyMean - xMean * yMean) / (x2Mean - xMean * xMean);
        const b = yMean - m * xMean;

        return [Math.exp(m), Math.exp(b)];
      }

      const ls = leastSquare(known_x, known_y);
      const m = ls[0];

      if(const_b){
        var b = ls[1];
      }
      else{
        var b = 1;
      }

      const result = [];

      for(var i = 0; i < new_x.length; i++){
        for(var j = 0; j < new_x[i].length; j++){
          const x = new_x[i][j];
          const y = b * Math.pow(m, x);
          // var y = Math.exp(b + m * x);

          result.push(Math.round(y * 1000000000) / 1000000000);
        }
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LINEST': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      return formula.error.v;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LOGEST': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      return formula.error.v;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
    'MDETERM': function() {
        // 参数个数校验：必须 1 个
        if (arguments.length !== 1) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 获取矩阵数据
            const data_array = arguments[0];
            let matrix = [];

            if (getObjType(data_array) === 'array') {
                if (getObjType(data_array[0]) === 'array' && !func_methods.isDyadicArr(data_array)) {
                    return formula.error.v;
                }
                matrix = func_methods.getDataDyadicArr(data_array);
            } else if (getObjType(data_array) === 'object' && data_array.startCell != null) {
                matrix = func_methods.getCellDataDyadicArr(data_array, 'number');
            } else {
                matrix = [[parseFloat(data_array)]];
            }

            // 校验所有值为数字
            const rows = matrix.length;
            for (let i = 0; i < rows; i++) {
                const cols = matrix[i].length;
                for (let j = 0; j < cols; j++) {
                    if (!isRealNum(matrix[i][j])) {
                        return formula.error.v;
                    }
                    matrix[i][j] = parseFloat(matrix[i][j]);
                }
            }

            // 必须是方阵
            const cols = matrix[0]?.length || 0;
            if (rows !== cols) {
                return formula.error.v;
            }

            // ====================== 通用行列式计算（N阶） ======================
            function determinant(mat) {
                const n = mat.length;

                // 1阶
                if (n === 1) {
                    return mat[0][0];
                }

                let det = 0;
                let sign = 1;

                for (let i = 0; i < n; i++) {
                    // 生成余子式
                    const subMatrix = [];
                    for (let row = 1; row < n; row++) {
                        const newRow = [];
                        for (let col = 0; col < n; col++) {
                            if (col !== i) {
                                newRow.push(mat[row][col]);
                            }
                        }
                        subMatrix.push(newRow);
                    }

                    det += mat[0][i] * sign * determinant(subMatrix);
                    sign = -sign;
                }

                return det;
            }

            return determinant(matrix);
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
  'MINVERSE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数组
      const data_array = arguments[0];
      let array = [];

      if(getObjType(data_array) == 'array'){
        if(getObjType(data_array[0]) == 'array' && !func_methods.isDyadicArr(data_array)){
          return formula.error.v;
        }

        array = func_methods.getDataDyadicArr(data_array);
      }
      else if(getObjType(data_array) == 'object' && data_array.startCell != null){
        array = func_methods.getCellDataDyadicArr(data_array, 'text');
      }
      else{
        const rowArr = [];
        rowArr.push(data_array);
        array.push(rowArr);
      }

      for(var i = 0; i < array.length; i++){
        for(let j = 0; j < array[i].length; j++){
          if(!isRealNum(array[i][j])){
            return formula.error.v;
          }

          array[i][j] = parseFloat(array[i][j]);
        }
      }

      if(array.length != array[0].length){
        return formula.error.v;
      }

      //计算
      return inverse(array);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'MMULT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      var p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数组1
      const data_array1 = arguments[0];
      let array1 = [];

      if(getObjType(data_array1) == 'array'){
        if(getObjType(data_array1[0]) == 'array' && !func_methods.isDyadicArr(data_array1)){
          return formula.error.v;
        }

        array1 = func_methods.getDataDyadicArr(data_array1);
      }
      else if(getObjType(data_array1) == 'object' && data_array1.startCell != null){
        array1 = func_methods.getCellDataDyadicArr(data_array1, 'text');
      }
      else{
        var rowArr = [];
        rowArr.push(data_array1);
        array1.push(rowArr);
      }

      for(var i = 0; i < array1.length; i++){
        for(var j = 0; j < array1[i].length; j++){
          if(!isRealNum(array1[i][j])){
            return formula.error.v;
          }

          array1[i][j] = parseFloat(array1[i][j]);
        }
      }

      //数组2
      const data_array2 = arguments[1];
      let array2 = [];

      if(getObjType(data_array2) == 'array'){
        if(getObjType(data_array2[0]) == 'array' && !func_methods.isDyadicArr(data_array2)){
          return formula.error.v;
        }

        array2 = func_methods.getDataDyadicArr(data_array2);
      }
      else if(getObjType(data_array2) == 'object' && data_array2.startCell != null){
        array2 = func_methods.getCellDataDyadicArr(data_array2, 'text');
      }
      else{
        var rowArr = [];
        rowArr.push(data_array2);
        array2.push(rowArr);
      }

      for(var i = 0; i < array2.length; i++){
        for(var j = 0; j < array2[i].length; j++){
          if(!isRealNum(array2[i][j])){
            return formula.error.v;
          }

          array2[i][j] = parseFloat(array2[i][j]);
        }
      }

      //计算
      if(array1[0].length != array2.length){
        return formula.error.v;
      }

      const rowlen = array1.length;
      const collen = array2[0].length;

      const result = [];

      for(let m = 0; m < rowlen; m++){
        var rowArr = [];

        for(let n = 0; n < collen; n++){
          let value = 0;

          for(var p = 0; p < array1[0].length; p++){
            value += array1[m][p] * array2[p][n];
          }

          rowArr.push(value);
        }

        result.push(rowArr);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SUMPRODUCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //第一个数组
      //数组1
      const data_array1 = arguments[0];
      let array1 = [];

      if(getObjType(data_array1) == 'array'){
        if(getObjType(data_array1[0]) == 'array' && !func_methods.isDyadicArr(data_array1)){
          return formula.error.v;
        }

        array1 = func_methods.getDataDyadicArr(data_array1);
      }
      else if(getObjType(data_array1) == 'object' && data_array1.startCell != null){
        array1 = func_methods.getCellDataDyadicArr(data_array1, 'text');
      }
      else{
        var rowArr = [];
        rowArr.push(data_array1);
        array1.push(rowArr);
      }

      for(var i = 0; i < array1.length; i++){
        for(let j = 0; j < array1[i].length; j++){
          if(!isRealNum(array1[i][j])){
            array1[i][j] = 0;
          }
          else{
            array1[i][j] = parseFloat(array1[i][j]);
          }
        }
      }

      const rowlen = array1.length;
      const collen = array1[0].length;

      if(arguments.length >= 2){
        for(var i = 1; i < arguments.length; i++){
          const data = arguments[i];
          let arr = [];

          if(getObjType(data) == 'array'){
            if(getObjType(data[0]) == 'array' && !func_methods.isDyadicArr(data)){
              return formula.error.v;
            }

            arr = func_methods.getDataDyadicArr(data);
          }
          else if(getObjType(data) == 'object' && data.startCell != null){
            arr = func_methods.getCellDataDyadicArr(data, 'text');
          }
          else{
            var rowArr = [];
            rowArr.push(data);
            arr.push(rowArr);
          }

          if(arr.length != rowlen || arr[0].length != collen){
            return formula.error.v;
          }

          for(var m = 0; m < rowlen; m++){
            for(var n = 0; n < collen; n++){
              if(!isRealNum(arr[m][n])){
                array1[m][n] = 0;
              }
              else{
                array1[m][n] = array1[m][n] * parseFloat(arr[m][n]);
              }
            }
          }
        }
      }

      let sum = 0;

      for(var m = 0; m < rowlen; m++){
        for(var n = 0; n < collen; n++){
          sum += array1[m][n];
        }
      }

      return sum;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISFORMULA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const data_cell = arguments[0];
      let cell;

      if(getObjType(data_cell) == 'object' && data_cell.startCell != null){
        if(data_cell.data == null){
          return false;
        }

        if(getObjType(data_cell.data) == 'array'){
          cell = data_cell.data[0][0];
        }
        else{
          cell = data_cell.data;
        }

        if(cell != null && cell.f != null){
          return true;
        }
        else{
          return false;
        }
      }
      else{
        return formula.error.v;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'CELL': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格信息的类型
      const data_info_type = arguments[0];
      let info_type;

      if(getObjType(data_info_type) == 'array'){
        if(getObjType(data_info_type[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_info_type)){
            return formula.error.v;
          }

          info_type = data_info_type[0][0];
        }
        else{
          info_type = data_info_type[0];
        }
      }
      else if(getObjType(data_info_type) == 'object' && data_info_type.startCell != null){
        if(data_info_type.data == null){
          return formula.error.v;
        }
        else{
          if(getObjType(data_info_type.data) == 'array'){
            return formula.error.v;
          }

          info_type = data_info_type.data.v;

          if(isRealNull(info_type)){
            return formula.error.v;
          }
        }
      }
      else{
        info_type = data_info_type;
      }

      //单元格
      const data_reference = arguments[1];
      let reference;

      if(getObjType(data_reference) == 'object' && data_reference.startCell != null){
        reference = data_reference.startCell;
      }
      else{
        return formula.error.v;
      }

      if(['address', 'col', 'color', 'contents', 'filename', 'format', 'parentheses', 'prefix', 'protect', 'row', 'type', 'width'].indexOf(info_type) == -1){
        return formula.error.v;
      }

      const file = getluckysheetfile()[getSheetIndex(Store.currentSheetIndex)];

      const cellrange = formula.getcellrange(reference);
      const row_index = cellrange.row[0];
      const col_index = cellrange.column[0];

      // let sheetdata = null;
      // sheetdata = Store.flowdata;
      // if (formula.execFunctionGroupData != null) {
      //     sheetdata = formula.execFunctionGroupData;
      // }

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      let value;
      if(formula.execFunctionGlobalData != null && formula.execFunctionGlobalData[`${row_index}_${col_index}_${Store.calculateSheetIndex}`]!=null){
        value = formula.execFunctionGlobalData[`${row_index}_${col_index}_${Store.calculateSheetIndex}`].v;
      }
      else if(sheetdata[row_index][col_index] != null && sheetdata[row_index][col_index].v != null && sheetdata[row_index][col_index].v !=''){
        value = sheetdata[row_index][col_index];
        if(value instanceof Object){
          value = value.v;
        }
      }
      else {
        value = 0;
      }

      switch(info_type){
      case 'address':
        return reference;
        break;
      case 'col':
        return col_index + 1;
        break;
      case 'color':
        return 0;
        break;
      case 'contents':
        // if (sheetdata[row_index][col_index] == null || sheetdata[row_index][col_index].v == null || sheetdata[row_index][col_index].v ==""){
        //     value = 0;
        // }

        return value;
        break;
      case 'filename':
        return file.name;
        break;
      case 'format':
        if (sheetdata[row_index][col_index] == null || sheetdata[row_index][col_index].ct == null){
          return 'G';
        }

        return sheetdata[row_index][col_index].ct.fa;
        break;
      case 'parentheses':
        if (sheetdata[row_index][col_index] == null || sheetdata[row_index][col_index].v == null || sheetdata[row_index][col_index].v ==''){
          return 0;
        }

        if (sheetdata[row_index][col_index].v > 0){
          return 1;
        }
        else{
          return 0;
        }
        break;
      case 'prefix':
        if (value==0){
          return '';
        }

        if (sheetdata[row_index][col_index].ht == 0){//居中对齐
          return '^';
        }
        else if (sheetdata[row_index][col_index].ht == 1){//左对齐
          return '\'';
        }
        else if (sheetdata[row_index][col_index].ht == 2){//右对齐
          return '"';
        }
        else{
          return '';
        }
        break;
      case 'protect':
        return 0;
        break;
      case 'row':
        return row_index + 1;
        break;
      case 'type':
        if (value==0){
          return 'b';
        }

        return 'l';
        break;
      case 'width':
        var cfg = file.config;

        if(cfg['columnlen'] != null && col_index in cfg['columnlen']){
          return cfg['columnlen'][col_index];
        }

        return Store.defaultcollen;
        break;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'NA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      return formula.error.na;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ERROR_TYPE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_error_val = arguments[0];
      let error_val;

      if(getObjType(data_error_val) == 'array'){
        if(getObjType(data_error_val[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_error_val)){
            return formula.error.v;
          }

          error_val = data_error_val[0][0];
        }
        else{
          error_val = data_error_val[0];
        }
      }
      else if(getObjType(data_error_val) == 'object' && data_error_val.startCell != null){
        if(data_error_val.data == null){
          return formula.error.na;
        }

        if(getObjType(data_error_val.data) == 'array'){
          error_val = data_error_val.data[0][0];

          if(error_val == null || isRealNull(error_val.v)){
            return formula.error.na;
          }

          error_val = error_val.v;
        }
        else{
          if(isRealNull(data_error_val.data.v)){
            return formula.error.na;
          }

          error_val = data_error_val.data.v;
        }
      }
      else{
        error_val = data_error_val;
      }

      const error_obj = {
        '#NULL!': 1,
        '#DIV/0!': 2,
        '#VALUE!': 3,
        '#REF!': 4,
        '#NAME?': 5,
        '#NUM!': 6,
        '#N/A': 7,
        '#GETTING_DATA': 8
      };

      if(error_val in error_obj){
        return error_obj[error_val];
      }
      else{
        return formula.error.na;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISBLANK': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_error_val = arguments[0];
      let error_val;

      if(getObjType(data_error_val) == 'object' && data_error_val.startCell != null){
        if(data_error_val.data == null){
          return true;
        }
        else{
          return false;
        }
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISERR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return true;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(['#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISERROR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return true;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISLOGICAL': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return false;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(value.toString().toLowerCase() == 'true' || value.toString().toLowerCase() == 'false'){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISNA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return false;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(value.toString() == '#N/A'){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISNONTEXT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return true;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return true;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return true;
      }
      else if(value.toString().toLowerCase() == 'true' || value.toString().toLowerCase() == 'false'){
        return true;
      }
      else if(isRealNum(value)){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISNUMBER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return false;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(isRealNum(value)){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISREF': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      if(getObjType(arguments[0]) == 'object' && arguments[0].startCell != null){
        return true;
      }
      else{
        return false;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'ISTEXT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return false;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return false;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return false;
      }
      else if(value.toString().toLowerCase() == 'true' || value.toString().toLowerCase() == 'false'){
        return false;
      }
      else if(isRealNum(value)){
        return false;
      }
      else{
        return true;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TYPE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        return 64;
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          return 16;
        }

        if(data_value.data == null || isRealNull(data_value.data.v)){
          return 1;
        }

        value = data_value.data.v;
      }
      else{
        value = data_value;
      }

      if(['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return 16;
      }
      else if(value.toString().toLowerCase() == 'true' || value.toString().toLowerCase() == 'false'){
        return 4;
      }
      else if(isRealNum(value)){
        return 1;
      }
      else{
        return 2;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'N': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //单元格
      const data_value = arguments[0];
      let value;

      if(getObjType(data_value) == 'array'){
        if(getObjType(data_value[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_value)){
            return formula.error.v;
          }

          value = data_value[0][0];
        }
        else{
          value = data_value[0];
        }
      }
      else if(getObjType(data_value) == 'object' && data_value.startCell != null){
        if(getObjType(data_value.data) == 'array'){
          value = data_value.data[0][0];

          if(value == null || isRealNull(value.v)){
            return 0;
          }

          value = value.v;
        }
        else{
          if(data_value.data == null || isRealNull(data_value.data.v)){
            return 0;
          }

          value = data_value.data.v;
        }
      }
      else{
        value = data_value;
      }

      if(['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].indexOf(value) > -1){
        return value;
      }
      else if(value.toString().toLowerCase() == 'true' || value.toString().toLowerCase() == 'false'){
        if(value.toString().toLowerCase() == 'true'){
          return 1;
        }
        else{
          return 0;
        }
      }
      else if(isRealNum(value)){
        return parseFloat(value);
      }
      else{
        return 0;
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TO_DATE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let value = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value)){
        return value;
      }

      if(!isRealNum(value)){
        return formula.error.v;
      }

      value = parseFloat(value);

      return update('yyyy-mm-dd', value);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TO_PURE_NUMBER': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const value = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(value)){
        return value;
      }

      if(dayjs(value).isValid()){
        return genarate(value)[2];
      }
      else{
        return numeral(value).value() == null ? value : numeral(value).value();
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TO_TEXT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const value = func_methods.getFirstValue(arguments[0], 'text');
      if(valueIsError(value)){
        return value;
      }

      return update('@', value);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TO_DOLLARS': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let value = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value)){
        return value;
      }

      if(!isRealNum(value)){
        return formula.error.v;
      }

      value = parseFloat(value);

      return update('$ 0.00', value);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TO_PERCENT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数字
      let value = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(value)){
        return value;
      }

      if(!isRealNum(value)){
        return formula.error.v;
      }

      value = parseFloat(value);

      return update('0%', value);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DGET': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      if (resultIndexes.length === 0) {
        return formula.error.v;
      }

      if (resultIndexes.length > 1) {
        return formula.error.nm;
      }

      return targetFields[resultIndexes[0]];
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DMAX': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let maxValue = targetFields[resultIndexes[0]];

      for (var i = 1; i < resultIndexes.length; i++) {
        if (maxValue < targetFields[resultIndexes[i]]) {
          maxValue = targetFields[resultIndexes[i]];
        }
      }

      return maxValue;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
    'DMIN': function() {
        // 参数个数校验：必须 3 个
        if (arguments.length !== 3) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // 1. 获取数据库区域
            const data_database = arguments[0];
            let database = [];

            if (getObjType(data_database) === 'object' && data_database.startCell != null) {
                if (!data_database.data) {
                    return formula.error.v;
                }
                database = func_methods.getCellDataDyadicArr(data_database, 'text');
            } else {
                return formula.error.v;
            }

            // 2. 获取字段（列）
            const field = func_methods.getFirstValue(arguments[1], 'text');
            if (valueIsError(field) || isRealNull(field)) {
                return formula.error.v;
            }

            // 3. 获取条件区域
            const data_criteria = arguments[2];
            let criteria = [];

            if (getObjType(data_criteria) === 'object' && data_criteria.startCell != null) {
                if (!data_criteria.data) {
                    return formula.error.v;
                }
                criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
            } else {
                return formula.error.v;
            }

            // 安全校验
            if (database.length < 2) return formula.error.v;
            if (criteria.length < 2) return formula.error.v;

            // ====================== 依赖函数（内置，避免外部依赖） ======================
            function findResultIndex(db, cri) {
                const headers = db[0];
                const criHeaders = cri[0];
                const res = [];

                for (let row = 1; row < db.length; row++) {
                    let match = true;
                    for (let ch = 0; ch < criHeaders.length; ch++) {
                        const h = criHeaders[ch];
                        const hidx = headers.indexOf(h);
                        if (hidx === -1) continue;

                        const cval = cri[1][ch];
                        const dval = db[row][hidx];

                        if (cval != null && cval !== '') {
                            if (String(dval) !== String(cval)) {
                                match = false;
                                break;
                            }
                        }
                    }
                    if (match) res.push(row);
                }
                return res;
            }

            function findField(db, f) {
                const h = db[0];
                const idx = h.findIndex(s => String(s).toUpperCase() === String(f).toUpperCase());
                return idx < 0 ? 0 : idx;
            }

            function rest(arr) {
                return arr.slice(1);
            }

            // ====================== 核心逻辑 ======================
            const resultIndexes = findResultIndex(database, criteria);
            if (resultIndexes.length === 0) return 0;

            let targetCol;
            if (isRealNum(field)) {
                const idx = parseInt(field, 10);
                targetCol = database[idx] || [];
            } else {
                const idx = findField(database, field);
                targetCol = database[idx] || [];
            }

            const values = rest(targetCol);
            let minValue = null;

            for (const i of resultIndexes) {
                const val = values[i - 1];
                if (!isRealNum(val)) continue;
                const num = parseFloat(val);
                if (minValue === null || num < minValue) minValue = num;
            }

            return minValue === null ? 0 : minValue;
        }
        catch (e) {
            return formula.errorInfo(e);
        }
    },
  'DAVERAGE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let sum = 0;

      for (var i = 0; i < resultIndexes.length; i++) {
        sum += targetFields[resultIndexes[i]];
      }

      return resultIndexes.length === 0 ? formula.error.d : sum / resultIndexes.length;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DCOUNT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      const targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      return window.luckysheet_function.COUNT.f.apply(window.luckysheet_function.COUNT, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DCOUNTA': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      const targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      return window.luckysheet_function.COUNTA.f.apply(window.luckysheet_function.COUNTA, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DPRODUCT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      let result = 1;

      for (i = 0; i < targetValues.length; i++) {
        result *= targetValues[i];
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DSTDEV': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      return window.luckysheet_function.STDEVA.f.apply(window.luckysheet_function.STDEVA, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DSTDEVP': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      return window.luckysheet_function.STDEVP.f.apply(window.luckysheet_function.STDEVP, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DSUM': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      let result = 0;

      for (i = 0; i < targetValues.length; i++) {
        result += targetValues[i];
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DVAR': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      return window.luckysheet_function.VAR_S.f.apply(window.luckysheet_function.VAR_S, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DVARP': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //数据库的单元格区域
      const data_database = arguments[0];
      let database = [];

      if(getObjType(data_database) == 'object' && data_database.startCell != null){
        if(data_database.data == null){
          return formula.error.v;
        }

        database = func_methods.getCellDataDyadicArr(data_database, 'text');
      }
      else{
        return formula.error.v;
      }

      //列
      const field = func_methods.getFirstValue(arguments[1], 'text');
      if(valueIsError(field)){
        return field;
      }

      if(isRealNull(field)){
        return formula.error.v;
      }

      //条件的单元格区域
      const data_criteria = arguments[2];
      let criteria = [];

      if(getObjType(data_criteria) == 'object' && data_criteria.startCell != null){
        if(data_criteria.data == null){
          return formula.error.v;
        }

        criteria = func_methods.getCellDataDyadicArr(data_criteria, 'text');
      }
      else{
        return formula.error.v;
      }

      if (!isRealNum(field) && getObjType(field) !== 'string') {
        return formula.error.v;
      }

      const resultIndexes = func_methods.findResultIndex(database, criteria);
      let targetFields = [];

      if (getObjType(field) === 'string') {
        const index = func_methods.findField(database, field);
        targetFields = func_methods.rest(database[index]);
      }
      else {
        targetFields = func_methods.rest(database[field]);
      }

      let targetValues = [];

      for (var i = 0; i < resultIndexes.length; i++) {
        targetValues[i] = targetFields[resultIndexes[i]];
      }

      targetValues = func_methods.compact(targetValues);

      return window.luckysheet_function.VAR_P.f.apply(window.luckysheet_function.VAR_P, targetValues);
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'LINESPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];
      let lineColor = arguments[1];
      let lineWidth = arguments[2];
      let normalValue = arguments[3];
      let normalColor = arguments[4];
      const maxSpot = arguments[5];
      const minSpot = arguments[6];
      let spotRadius = arguments[7];

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      if(lineWidth==null){
        lineWidth = 1;
      }
      sparksetting['lineWidth'] = lineWidth;
      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting['offsetY'] = lineWidth+1;
      sparksetting.height = height-(lineWidth+1);
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      if(lineColor==null){
        lineColor = '#2ec7c9';
      }
      sparkColorSetting('lineColor', lineColor);
      //sparkColorSetting("fillColor", fillColor);
      sparksetting['fillColor'] = 0;


      //设置辅助线，可以支持min、max、avg、median等几个字符变量，或者具体的数值。
      if(normalValue){
        if(typeof(normalValue)==='string'){
          normalValue = normalValue.toLowerCase();
          let nv = null;
          if(normalValue=='min'){
            nv = window.luckysheet_function.MIN.f({'data':dataformat});
          }
          else if(normalValue=='max'){
            nv = window.luckysheet_function.MAX.f({'data':dataformat});
          }
          else if(normalValue=='avg' || normalValue=='mean'){
            nv = window.luckysheet_function.AVERAGE.f({'data':dataformat});
          }
          else if(normalValue=='median'){
            nv = window.luckysheet_function.MEDIAN.f({'data':dataformat});
          }

          if(nv){
            sparksetting['normalRangeMin'] = nv;
            sparksetting['normalRangeMax'] = nv;
          }
        }
        else{
          sparksetting['normalRangeMin'] = normalValue;
          sparksetting['normalRangeMax'] = normalValue;
        }

      }

      if(normalColor==null){
        normalColor = '#000';
      }
      sparkColorSetting('normalRangeColor', normalColor);

      sparkColorSetting('maxSpotColor', maxSpot);
      sparkColorSetting('minSpotColor', minSpot);

      if(spotRadius==null){
        spotRadius = '1.5';
      }
      sparksetting['spotRadius'] = spotRadius;


      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
      // {
      //     height:rowlen,
      //     width:firstcolumnlen,
      //     normalRangeMin:6,
      //     normalRangeMax:6,
      //     normalRangeColor:"#000"
      // }
      //return "";
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'AREASPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];
      let lineColor = arguments[1];
      const fillColor = arguments[2];
      let lineWidth = arguments[3];
      let normalValue = arguments[4];
      let normalColor = arguments[5];
      // var maxSpot = arguments[5];
      // var minSpot = arguments[6];
      // var spotRadius = arguments[7];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      if(lineWidth==null){
        lineWidth = 1;
      }
      sparksetting['lineWidth'] = lineWidth;
      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting['offsetY'] = lineWidth+1;
      sparksetting.height = height-(lineWidth+1);
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      if(lineColor==null){
        lineColor = '#2ec7c9';
      }
      sparkColorSetting('lineColor', lineColor);
      sparkColorSetting('fillColor', fillColor);
      // sparksetting["fillColor"] = 0;

      if(lineWidth==null){
        lineWidth = '1';
      }
      sparksetting['lineWidth'] = lineWidth;

      //设置辅助线，可以支持min、max、avg、median等几个字符变量，或者具体的数值。
      if(normalValue){
        if(typeof(normalValue)==='string'){
          normalValue = normalValue.toLowerCase();
          let nv = null;
          if(normalValue=='min'){
            nv = window.luckysheet_function.MIN.f({'data':dataformat});
          }
          else if(normalValue=='max'){
            nv = window.luckysheet_function.MAX.f({'data':dataformat});
          }
          else if(normalValue=='avg' || normalValue=='mean'){
            nv = window.luckysheet_function.AVERAGE.f({'data':dataformat});
          }
          else if(normalValue=='median'){
            nv = window.luckysheet_function.MEDIAN.f({'data':dataformat});
          }

          if(nv){
            sparksetting['normalRangeMin'] = nv;
            sparksetting['normalRangeMax'] = nv;
          }
        }
        else{
          sparksetting['normalRangeMin'] = normalValue;
          sparksetting['normalRangeMax'] = normalValue;
        }

      }

      if(normalColor==null){
        normalColor = '#000';
      }
      sparkColorSetting('normalRangeColor', normalColor);

      // sparkColorSetting("maxSpotColor", maxSpot);
      // sparkColorSetting("minSpotColor", minSpot);

      // if(spotRadius==null){
      //     spotRadius = "1.5";
      // }
      // sparksetting["spotRadius"] = spotRadius;


      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
      // {
      //     height:rowlen,
      //     width:firstcolumnlen,
      //     normalRangeMin:6,
      //     normalRangeMax:6,
      //     normalRangeColor:"#000"
      // }
      //return "";
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'COLUMNSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let barSpacing = arguments[1];
      let barColor = arguments[2];
      let negBarColor = arguments[3];
      const chartRangeMax = arguments[4];

      ////具体实现
      sparksetting['type'] = 'column';
      if(barSpacing==null){
        barSpacing = '1';
      }
      sparksetting['barSpacing'] = barSpacing;

      if(barColor==null){
        barColor = '#fc5c5c';
      }
      sparkColorSetting('barColor', barColor);

      if(negBarColor==null){
        negBarColor = '#97b552';
      }
      sparkColorSetting('negBarColor', negBarColor);

      if(chartRangeMax==null || chartRangeMax===false || typeof chartRangeMax !=='number' ){
        sparksetting['chartRangeMax'] = undefined;
      }
      else{
        sparksetting['chartRangeMax'] = chartRangeMax;
      }

      const colorLists = formula.sparklinesColorMap(arguments);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'STACKCOLUMNSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      //var dataformat = formula.readCellDataToOneArray(rangeValue);

      const dataformat = [];

      let data = [];
      if(rangeValue!=null && rangeValue.data!=null){
        data = rangeValue.data;
      }

      if(getObjType(data) == 'array'){
        data = formula.getPureValueByData(data);
      }
      else if(getObjType(data) == 'object'){
        data = data.v;

        return [data];
      }
      else{
        if(/\{.*?\}/.test(data)){
          data = data.replace(/\{/g, '[').replace(/\}/g, ']');
        }
        data = new Function(`return ${  data}`)();
      }

      const stackconfig = arguments[1];
      var offsetY = data.length;
      if(stackconfig==null || !!stackconfig){
        for(var c=0;c<data[0].length;c++){
          let colstr = '';
          for(var r=0;r<data.length;r++){
            colstr += `${data[r][c]  }:`;
          }
          colstr = colstr.substr(0, colstr.length-1);
          dataformat.push(colstr);
        }
      }
      else{
        for(var r=0;r<data.length;r++){
          let rowstr = '';
          for(var c=0;c<data[0].length;c++){
            rowstr += `${data[r][c]  }:`;
          }
          rowstr = rowstr.substr(0, rowstr.length-1);
          dataformat.push(rowstr);
        }
        var offsetY = data[0].length;
      }

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;
      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;
      //sparksetting.offsetY = offsetY;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let barSpacing = arguments[2];
      const chartRangeMax = arguments[3];

      ////具体实现
      sparksetting['type'] = 'column';
      if(barSpacing==null){
        barSpacing = '1';
      }
      sparksetting['barSpacing'] = barSpacing;

      if(chartRangeMax==null || chartRangeMax===false || typeof chartRangeMax !=='number' ){
        sparksetting['chartRangeMax'] = undefined;
      }
      else{
        sparksetting['chartRangeMax'] = chartRangeMax;
      }

      const colorLists = formula.sparklinesColorMap(arguments, 4);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现


      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BARSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let barSpacing = arguments[1];
      let barColor = arguments[2];
      let negBarColor = arguments[3];
      const chartRangeMax = arguments[4];

      ////具体实现
      sparksetting['type'] = 'bar';
      if(barSpacing==null){
        barSpacing = '1';
      }
      sparksetting['barSpacing'] = barSpacing;

      if(barColor==null){
        barColor = '#fc5c5c';
      }
      sparkColorSetting('barColor', barColor);

      if(negBarColor==null){
        negBarColor = '#97b552';
      }
      sparkColorSetting('negBarColor', negBarColor);

      if(chartRangeMax==null || chartRangeMax===false || typeof chartRangeMax !=='number' ){
        sparksetting['chartRangeMax'] = undefined;
      }
      else{
        sparksetting['chartRangeMax'] = chartRangeMax;
      }

      const colorLists = formula.sparklinesColorMap(arguments);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'STACKBARSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      //var dataformat = formula.readCellDataToOneArray(rangeValue);

      const dataformat = [];

      let data = [];
      if(rangeValue!=null && rangeValue.data!=null){
        data = rangeValue.data;
      }

      if(getObjType(data) == 'array'){
        data = formula.getPureValueByData(data);
      }
      else if(getObjType(data) == 'object'){
        data = data.v;
        return [data];
      }
      else{
        if(/\{.*?\}/.test(data)){
          data = data.replace(/\{/g, '[').replace(/\}/g, ']');
        }
        data = new Function(`return ${  data}`)();
      }

      const stackconfig = arguments[1];
      var offsetY = data.length;
      if(stackconfig==null || !!stackconfig){
        for(var c=0;c<data[0].length;c++){
          let colstr = '';
          for(var r=0;r<data.length;r++){
            colstr += `${data[r][c]  }:`;
          }
          colstr = colstr.substr(0, colstr.length-1);
          dataformat.push(colstr);
        }
      }
      else{
        for(var r=0;r<data.length;r++){
          let rowstr = '';
          for(var c=0;c<data[0].length;c++){
            rowstr += `${data[r][c]  }:`;
          }
          rowstr = rowstr.substr(0, rowstr.length-1);
          dataformat.push(rowstr);
        }
        var offsetY = data[0].length;
      }

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;
      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;
      //sparksetting.offsetY = offsetY;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let barSpacing = arguments[2];
      const chartRangeMax = arguments[3];

      ////具体实现
      sparksetting['type'] = 'bar';
      if(barSpacing==null){
        barSpacing = '1';
      }
      sparksetting['barSpacing'] = barSpacing;

      if(chartRangeMax==null || chartRangeMax===false || typeof chartRangeMax !=='number' ){
        sparksetting['chartRangeMax'] = undefined;
      }
      else{
        sparksetting['chartRangeMax'] = chartRangeMax;
      }

      const colorLists = formula.sparklinesColorMap(arguments, 4);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现


      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'DISCRETESPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let thresholdValue = arguments[1];
      let barColor = arguments[2];
      let negBarColor = arguments[3];

      ////具体实现
      sparksetting['type'] = 'discrete';

      if(thresholdValue==null){
        thresholdValue = 0;
      }
      sparksetting['thresholdValue'] = thresholdValue;

      if(barColor==null){
        barColor = '#2ec7c9';
      }
      sparkColorSetting('lineColor', barColor);

      if(negBarColor==null){
        negBarColor = '#fc5c5c';
      }
      sparkColorSetting('thresholdColor', negBarColor);
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'TRISTATESPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let barSpacing = arguments[1];
      let barColor = arguments[2];
      let negBarColor = arguments[3];
      let zeroBarColor = arguments[4];

      ////具体实现
      sparksetting['type'] = 'tristate';
      if(barSpacing==null){
        barSpacing = '1';
      }
      sparksetting['barSpacing'] = barSpacing;

      if(barColor==null){
        barColor = '#fc5c5c';
      }
      sparkColorSetting('barColor', barColor);

      if(negBarColor==null){
        negBarColor = '#97b552';
      }
      sparkColorSetting('negBarColor', negBarColor);

      if(zeroBarColor==null){
        zeroBarColor = '#999';
      }
      sparkColorSetting('zeroBarColor', zeroBarColor);

      const colorLists = formula.sparklinesColorMap(arguments);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'PIESPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let offset = arguments[1];
      let borderWidth = arguments[2];
      let borderColor = arguments[3];

      ////具体实现
      sparksetting['type'] = 'pie';
      if(offset==null){
        offset = 0;
      }
      sparksetting['offset'] = offset;

      if(borderWidth==null){
        borderWidth = 0;
      }
      sparkColorSetting('borderWidth', borderWidth);

      if(borderColor==null){
        borderColor = '#97b552';
      }
      sparkColorSetting('borderColor', borderColor);

      const colorLists = formula.sparklinesColorMap(arguments, 4);
      if(colorLists){
        sparksetting['colorMap'] = colorLists;
      }
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BOXSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来���用到
      const colorList = formula.colorList;
      const rangeValue = arguments[0];

      //定义需要格式化data数据
      const dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      let outlierIQR = arguments[1];
      let target = arguments[2];
      let spotRadius = arguments[3];

      ////具体实现
      sparksetting['type'] = 'box';
      if(outlierIQR==null){
        outlierIQR = 1.5;
      }
      sparksetting['outlierIQR'] = outlierIQR;

      if(target==null){
        target = 0;
      }
      else{
        sparkColorSetting('target', target);
      }

      if(spotRadius==null){
        spotRadius = 1.5;
      }
      sparkColorSetting('spotRadius', spotRadius);
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'BULLETSPLINES': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const cell_fp = window.luckysheetCurrentFunction;
      //色表，接下来会用到
      const colorList = formula.colorList;
      //var rangeValue = arguments[0];

      //定义需要格式化data数据
      //var dataformat = formula.readCellDataToOneArray(rangeValue);

      const luckysheetfile = getluckysheetfile();
      const index = getSheetIndex(Store.calculateSheetIndex);
      const sheetdata = luckysheetfile[index].data;

      //在下面获得该单元格的长度和宽度,同时考虑了合并单元格问题
      const cellSize = menuButton.getCellRealSize(sheetdata, cell_r, cell_c);
      const width = cellSize[0];
      const height = cellSize[1];

      //开始进行sparklines的详细设置，宽和高为单元格的宽高。
      const sparksetting = {};

      //设置sparklines图表的宽高，线图的高会随着粗细而超出单元格高度，所以减去一个量，设置offsetY或者offsetX为渲染偏移量，传给luckysheetDrawMain使用。默认为0。=LINESPLINES(D9:E24,3,5)
      sparksetting.height = height;
      sparksetting.width = width;

      //定义sparklines的通用色彩设置函数，可以设置 色表【colorList】索引数值 或者 具体颜色值
      const sparkColorSetting = function(attr, value){
        if(value){
          if(typeof(value)==='number'){
            if(value>19){
              value = value % 20;
            }
            value = colorList[value];
          }
          sparksetting[attr] = value;
        }
      };

      ////具体实现
      const dataformat = [];
      luckysheet_getValue(arguments);

      const data1 = formula.getValueByFuncData(arguments[0]);
      const data2 = formula.getValueByFuncData(arguments[1]);

      dataformat.push(data1);
      dataformat.push(data2);

      for(var i=2;i<arguments.length;i++){
        dataformat.push(formula.getValueByFuncData(arguments[i]));
      }

      sparksetting['type'] = 'bullet';
      ////具体实现

      const temp1 = luckysheetSparkline.init(dataformat, sparksetting);

      return temp1;
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  //动态数组公式
  'SORT': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要排序的范围或数组
      const data_array = arguments[0];
      let array = [], rowlen = 1, collen = 1;

      if(getObjType(data_array) == 'array'){
        if(getObjType(data_array[0]) == 'array'){
          if(!func_methods.isDyadicArr(data_array)){
            return formula.error.v;
          }

          for(var i = 0; i < data_array.length; i++){
            var rowArr = [];

            for(var j = 0; j < data_array[i].length; j++){
              var number = data_array[i][j];

              rowArr.push(number);
            }

            array.push(rowArr);
          }

          rowlen = array.length;
          collen = array[0].length;
        }
        else{
          for(var i = 0; i < data_array.length; i++){
            var number = data_array[i];

            array.push(number);
          }

          rowlen = array.length;
        }
      }
      else if(getObjType(data_array) == 'object' && data_array.startCell != null){
        if(data_array.data != null){
          if(getObjType(data_array.data) == 'array'){
            for(var i = 0; i < data_array.data.length; i++){
              var rowArr = [];

              for(var j = 0; j < data_array.data[i].length; j++){
                if(data_array.data[i][j] != null){
                  var number = data_array.data[i][j].v;

                  if(isRealNull(number)){
                    number = 0;
                  }

                  rowArr.push(number);
                }
                else{
                  rowArr.push(0);
                }
              }

              array.push(rowArr);
            }

            rowlen = array.length;
            collen = array[0].length;
          }
          else{
            var number = data_array.data.v;

            if(isRealNull(number)){
              number = 0;
            }

            array.push(number);
          }
        }
        else{
          array.push(0);
        }
      }
      else{
        var number = data_array;

        array.push(number);
      }

      //表示要排序的行或列的数字（默认row1/col1）
      let sort_index = 1;
      if(arguments.length >= 2){
        sort_index = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(sort_index)){
          return sort_index;
        }

        if(!isRealNum(sort_index)){
          return formula.error.v;
        }

        sort_index = parseInt(sort_index);
      }

      //表示所需排序顺序的数字；1表示升序（默认），-1表示降序。
      let sort_order = 1;
      if(arguments.length >= 3){
        sort_order = func_methods.getFirstValue(arguments[2]);
        if(valueIsError(sort_order)){
          return sort_order;
        }

        if(!isRealNum(sort_order)){
          return formula.error.v;
        }

        sort_order = Math.floor(parseFloat(sort_order));
      }

      //表示所需排序方向的逻辑值；按行排序为FALSE（默认），按列排序为TRUE。
      let by_col = false;
      if(arguments.length == 4){
        by_col = func_methods.getCellBoolen(arguments[3]);

        if(valueIsError(by_col)){
          return by_col;
        }
      }

      if(by_col){
        if(sort_index < 1 || sort_index > rowlen){
          return formula.error.v;
        }
      }
      else{
        if(sort_index < 1 || sort_index > collen){
          return formula.error.v;
        }
      }

      if(sort_order != 1 && sort_order != -1){
        return formula.error.v;
      }

      //计算
      const asc = function(x, y){
        if(getObjType(x) == 'array'){
          x = x[sort_index - 1];
        }

        if(getObjType(y) == 'array'){
          y = y[sort_index - 1];
        }

        if(!isNaN(x) && !isNaN(y)){
          return x - y;
        }
        else if(!isNaN(x)){
          return -1;
        }
        else if(!isNaN(y)){
          return 1;
        }
        else{
          if(x > y){
            return 1;
          }
          else if(x < y){
            return -1;
          }
        }
      };

      const desc = function(x, y){
        if(getObjType(x) == 'array'){
          x = x[sort_index - 1];
        }

        if(getObjType(y) == 'array'){
          y = y[sort_index - 1];
        }

        if(!isNaN(x) && !isNaN(y)){
          return y - x;
        }
        else if(!isNaN(x)){
          return 1;
        }
        else if(!isNaN(y)){
          return -1;
        }
        else{
          if(x > y){
            return -1;
          }
          else if(x < y){
            return 1;
          }
        }
      };

      if(by_col){
        array = array[0].map(function(col, a){
          return array.map(function(row){
            return row[a];
          });
        });

        if(sort_order == 1){
          array.sort(asc);
        }

        if(sort_order == -1){
          array.sort(desc);
        }

        array = array[0].map(function(col, b){
          return array.map(function(row){
            return row[b];
          });
        });
      }
      else{
        if(sort_order == 1){
          array.sort(asc);
        }

        if(sort_order == -1){
          array.sort(desc);
        }
      }

      return array;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
    'FILTER': function() {
        // 参数个数校验：2~3个参数
        if (arguments.length < 2 || arguments.length > 3) {
            return formula.error.na;
        }

        // 参数类型校验
        for (let i = 0; i < arguments.length; i++) {
            const p = formula.errorParamCheck(this.p, arguments[i], i);
            if (!p[0]) {
                return formula.error.v;
            }
        }

        try {
            // ==============================================
            // 1. 获取要筛选的数组
            // ==============================================
            const data_array = arguments[0];
            let array = [];

            if (getObjType(data_array) === 'array') {
                if (getObjType(data_array[0]) === 'array' && !func_methods.isDyadicArr(data_array)) {
                    return formula.error.v;
                }
                array = func_methods.getDataDyadicArr(data_array);
            } else if (getObjType(data_array) === 'object' && data_array.startCell != null) {
                array = func_methods.getCellDataDyadicArr(data_array, 'text');
            } else {
                array = [[data_array]];
            }

            if (!array || !array.length || !array[0] || !array[0].length) {
                return formula.error.na;
            }

            const rowLen = array.length;
            const colLen = array[0].length;

            // ==============================================
            // 2. 获取并解析布尔条件数组
            // ==============================================
            const data_include = arguments[1];
            let include = [];
            let filterType = 'row'; // row：按行筛选 | col：按列筛选

            // 统一提取条件数据
            let includeArr = [];
            if (getObjType(data_include) === 'object' && data_include.data) {
                includeArr = func_methods.getCellDataDyadicArr(data_include, 'text');
            } else {
                includeArr = data_include;
            }

            // 条件必须是一维数组（单行或单列）
            if (!Array.isArray(includeArr) || !includeArr.length) {
                return formula.error.v;
            }

            const isRowCondition = includeArr.length > 1;
            const isColCondition = !isRowCondition && includeArr[0]?.length > 1;

            // 行筛选（垂直条件）
            if (isRowCondition) {
                if (includeArr.length !== rowLen) return formula.error.v;
                filterType = 'row';
                include = includeArr.map(row => toBoolean(row[0]));
            }
            // 列筛选（水平条件）
            else if (isColCondition) {
                if (includeArr[0].length !== colLen) return formula.error.v;
                filterType = 'col';
                include = includeArr[0].map(val => toBoolean(val));
            } else {
                return formula.error.v;
            }

            // ==============================================
            // 3. 空值返回内容
            // ==============================================
            let if_empty = formula.error.na;
            if (arguments.length === 3) {
                if_empty = func_methods.getFirstValue(arguments[2], 'text');
            }

            // ==============================================
            // 4. 执行筛选
            // ==============================================
            const result = [];

            // 按行筛选
            if (filterType === 'row') {
                for (let i = 0; i < rowLen; i++) {
                    if (include[i]) {
                        result.push([...array[i]]);
                    }
                }
            }
            // 按列筛选
            else {
                for (let i = 0; i < rowLen; i++) {
                    const newRow = [];
                    for (let j = 0; j < colLen; j++) {
                        if (include[j]) {
                            newRow.push(array[i][j]);
                        }
                    }
                    if (newRow.length > 0) {
                        result.push(newRow);
                    }
                }
            }

            // 无结果时返回指定内容
            return result.length === 0 ? if_empty : result;
        }
        catch (e) {
            return formula.errorInfo(e);
        }

        // 内部工具：任意值转布尔
        function toBoolean(val) {
            if (typeof val === 'boolean') return val;
            if (isRealNum(val)) return parseFloat(val) !== 0;
            if (typeof val === 'string') {
                const lower = val.trim().toLowerCase();
                if (lower === 'true') return true;
                if (lower === 'false') return false;
            }
            return false;
        }
    },
  'UNIQUE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //从其返回唯一值的数组或区域
      const data_array = arguments[0];
      let array = [];

      if(getObjType(data_array) == 'array'){
        if(getObjType(data_array[0]) == 'array' && !func_methods.isDyadicArr(data_array)){
          return formula.error.v;
        }

        array = func_methods.getDataDyadicArr(data_array);
      }
      else if(getObjType(data_array) == 'object' && data_array.startCell != null){
        array = func_methods.getCellDataDyadicArr(data_array, 'number');
      }
      else{
        const rowArr = [];

        rowArr.push(parseFloat(data_array));

        array.push(rowArr);
      }

      //逻辑值，指示如何比较；按行 = FALSE 或省略；按列 = TRUE
      let by_col = false;
      if(arguments.length >= 2){
        by_col = func_methods.getCellBoolen(arguments[1]);

        if(valueIsError(by_col)){
          return by_col;
        }
      }

      //逻辑值，仅返回唯一值中出现一次 = TRUE；包括所有唯一值 = FALSE 或省略
      let occurs_once = false;
      if(arguments.length == 3){
        occurs_once = func_methods.getCellBoolen(arguments[2]);

        if(valueIsError(occurs_once)){
          return occurs_once;
        }
      }

      //计算
      if(by_col){
        array = array[0].map(function(col, a){
          return array.map(function(row){
            return row[a];
          });
        });

        var strObj = {}, strArr = [];
        var allUnique = [];

        for(var i = 0; i < array.length; i++){
          var str = '';

          for(var j = 0; j < array[i].length; j++){
            str += `${array[i][j].toString()  }|||`;
          }

          strArr.push(str);

          if(!(str in strObj)){
            strObj[str] = 0;

            allUnique.push(array[i]);
          }
        }

        if(occurs_once){
          var oneUnique = [];

          for(var i = 0; i < strArr.length; i++){
            if(strArr.indexOf(strArr[i]) == strArr.lastIndexOf(strArr[i])){
              oneUnique.push(array[i]);
            }
          }

          oneUnique = oneUnique[0].map(function(col, a){
            return oneUnique.map(function(row){
              return row[a];
            });
          });

          return oneUnique;
        }
        else{
          allUnique = allUnique[0].map(function(col, a){
            return allUnique.map(function(row){
              return row[a];
            });
          });

          return allUnique;
        }
      }
      else{
        var strObj = {}, strArr = [];
        var allUnique = [];

        for(var i = 0; i < array.length; i++){
          var str = '';

          for(var j = 0; j < array[i].length; j++){
            str += `${array[i][j].toString()  }|||`;
          }

          strArr.push(str);

          if(!(str in strObj)){
            strObj[str] = 0;

            allUnique.push(array[i]);
          }
        }

        if(occurs_once){
          var oneUnique = [];

          for(var i = 0; i < strArr.length; i++){
            if(strArr.indexOf(strArr[i]) == strArr.lastIndexOf(strArr[i])){
              oneUnique.push(array[i]);
            }
          }

          return oneUnique;
        }
        else{
          return allUnique;
        }
      }
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'RANDARRAY': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要返回的行数
      let rows = 1;
      if(arguments.length >= 1){
        rows = func_methods.getFirstValue(arguments[0]);
        if(valueIsError(rows)){
          return rows;
        }

        if(!isRealNum(rows)){
          return formula.error.v;
        }

        rows = parseInt(rows);
      }

      //要返回的列数
      let cols = 1;
      if(arguments.length == 2){
        cols = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(cols)){
          return cols;
        }

        if(!isRealNum(cols)){
          return formula.error.v;
        }

        cols = parseInt(cols);
      }

      if(rows <= 0 || cols <= 0){
        return formula.error.v;
      }

      //计算
      const result = [];

      for(var i = 0; i < rows; i++){
        const result_row = [];

        for(let j = 0; j < cols; j++){
          result_row.push(Math.random().toFixed(9));
        }

        result.push(result_row);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'SEQUENCE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (var i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      //要返回的行数
      let rows = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(rows)){
        return rows;
      }

      if(!isRealNum(rows)){
        return formula.error.v;
      }

      rows = parseInt(rows);

      //要返回的列数
      let cols = 1;
      if(arguments.length >= 2){
        cols = func_methods.getFirstValue(arguments[1]);
        if(valueIsError(cols)){
          return cols;
        }

        if(!isRealNum(cols)){
          return formula.error.v;
        }

        cols = parseInt(cols);
      }

      //序列中的第一个数字
      let start = 1;
      if(arguments.length >= 3){
        start = func_methods.getFirstValue(arguments[2]);
        if(valueIsError(start)){
          return start;
        }

        if(!isRealNum(start)){
          return formula.error.v;
        }

        start = parseFloat(start);
      }

      //序列中每个序列值的增量
      let step = 1;
      if(arguments.length == 4){
        step = func_methods.getFirstValue(arguments[3]);
        if(valueIsError(step)){
          return step;
        }

        if(!isRealNum(step)){
          return formula.error.v;
        }

        step = parseFloat(step);
      }

      if(rows <= 0 || cols <= 0){
        return formula.error.v;
      }

      //计算
      const result = [];

      for(var i = 0; i < rows; i++){
        const result_row = [];

        for(let j = 0; j < cols; j++){
          const number = start + step * (j + cols * i);
          result_row.push(number);
        }

        result.push(result_row);
      }

      return result;
    }
    catch (e) {
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'EVALUATE': function() {
    //必要参数个数错误检测
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    //参数类型错误检测
    for (let i = 0; i < arguments.length; i++) {
      const p = formula.errorParamCheck(this.p, arguments[i], i);

      if (!p[0]) {
        return formula.error.v;
      }
    }

    try {
      const cell_r = window.luckysheetCurrentRow;
      const cell_c = window.luckysheetCurrentColumn;
      const sheetindex_now = window.luckysheetCurrentIndex;
      //公式文本
      let strtext = func_methods.getFirstValue(arguments[0]).toString();
      if(valueIsError(strtext)){
        return strtext;
      }
      //在文本公式前面添加=
      if(strtext.trim().indexOf('=')!=0)
      {
        strtext =`=${strtext}`;
      }
      //console.log(strtext);
      const result_this = formula.execstringformula(strtext,cell_r,cell_c,sheetindex_now);
      return result_this[1];
    }
    catch (e) {
      let err = e;
      //计算错误检测
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
  'REMOTE': function() {
    if (arguments.length < this.m[0] || arguments.length > this.m[1]) {
      return formula.error.na;
    }

    try {
      const cellRow = window.luckysheetCurrentRow;
      const cellColumn = window.luckysheetCurrentColumn;
      const cellFunction = window.luckysheetCurrentFunction;

      const remoteFunction = func_methods.getFirstValue(arguments[0]);
      if(valueIsError(remoteFunction)){
        return remoteFunction;
      }

      luckysheetConfigsetting.remoteFunction(remoteFunction, data => {
        const flowData = editor.deepCopyFlowData(Store.flowdata);
        formula.execFunctionGroup(cellRow, cellColumn, data);
        flowData[cellRow][cellColumn] = {
          'v': data,
          'f': cellFunction
        };
        jfrefreshgrid(flowData, [{'row': [cellRow, cellRow], 'column': [cellColumn, cellColumn]}]);
      });

      return 'Loading...';
    }
    catch (e) {
      console.log(e);
      let err = e;
      err = formula.errorInfo(err);
      return [formula.error.v, err];
    }
  },
};

export default functionImplementation;