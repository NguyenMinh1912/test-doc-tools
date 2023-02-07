const xlsx = require('xlsx');

String.prototype.format = function () {
    // store arguments in an array
    var args = arguments;
    // use replace to iterate over the string
    // select the match and check if the related argument is present
    // if yes, replace the match with the argument
    return this.replace(/{([0-9]+)}/g, function (match, index) {
      // check if the argument is present
      return typeof args[index] == 'undefined' ? match : args[index];
    });
};

const descriptionBuilder = () => {
    if (methodTestName.includes('Error')) {
        return `（異常系）
        Character type : {0}`;
    }
    return `（正常系）
    Character type
    {0}`;
}

const filePath = 'テストコード「/tr-cccpf-cvos-common/src/test/java/jp/co/dnp/cio/ngsp/app/common/logic/UpdateAppBrandImplFuncTest{0}」参照';
const type = 'ParameterCheck';

const methodTestName = 'validateError11';
const xlsxData = [];

const generateTestDoc = (values, testTarget, methodTestName) => {
    xlsxData.push(...values.map(value => {
       return {
           'NO': '',
           '重要度': 'A',
           'テストNo': '',
           '機能中項目': type,
           '機能小項目': testTarget,
           'テスト観点': descriptionBuilder().format(value),
           '実施条件２': '',
           '実施条件３': '',
           '実施手順': filePath.format(`.${methodTestName}()`),
           '入力値':  filePath.format(`.java`),
           '期待結果':  filePath.format(`.java`),
           '結果':'',
           '実施日':'',
           '実施者':'',
           'バグID':'',
           'バージョン':'',
           'リグレッションテスト':'',
           '実施日':'',
           '実施者':'',
           'バグID':'',
           'バージョン':'',
           'リグレッションテスト':'',
           '実施日':'',
           '実施者':'',
           'バグID':'',
           'バージョン':'',
           'リグレッションテスト':'',
           '備考':'',
       }
   }))

}

generateTestDoc(["", "0", "1234", "MAX_VALUE", 'FULL_SIZE_NUMBER', 'KANATA_HALFSIZE',
'BREAK_LINE', 'ALPHABET_FULLSIZE', 'NUMBER_FULLSIZE', 'KANATA_FULLSIZE',
'KANJI_FULLSIZE', 'HIRAGANA_FULLSIZE', 'SYMBOL_FULLSIZE', 'SPACE_FULLSIZE', 'SPACE_HALFSIZE',
'SPACE_BETWEEN_HALFSIZE', 'SPACE_BETWEEN_FULLSIZE', 'SYMBOL_HALFSIZE_1', 'CHAR_HALFSIZE'],
'version',
'validateError11'
)

generateTestDoc(["", "0", "1234", "MAX_VALUE", 'FULL_SIZE_NUMBER', 'KANATA_HALFSIZE',
'BREAK_LINE', 'ALPHABET_FULLSIZE', 'NUMBER_FULLSIZE', 'KANATA_FULLSIZE',
'KANJI_FULLSIZE', 'HIRAGANA_FULLSIZE', 'SYMBOL_FULLSIZE', 'SPACE_FULLSIZE', 'SPACE_HALFSIZE',
'SPACE_BETWEEN_HALFSIZE', 'SPACE_BETWEEN_FULLSIZE', 'SYMBOL_HALFSIZE_1', 'CHAR_HALFSIZE'],
'version',
'validateError12'
)

const wb = xlsx.utils.book_new();
    
const sheet = xlsx.utils.json_to_sheet(xlsxData, {});

xlsx.utils.book_append_sheet(wb, sheet);

xlsx.writeFile(wb, 'test.xlsx');





