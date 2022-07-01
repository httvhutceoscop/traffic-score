$(document).ready(function () {
  const MSG_MAXIMUM_FILES = "Maximum files for upload are 6 files. Please decrease your files.";
  const KEY = {
    "LAST_NAME": "チャプター► サマリーPALMSレポート",
    "FIRST_NAME": "__EMPTY",
    "ATTENDANCE": "__EMPTY_1",
    "ABSENT": "__EMPTY_2",
    "LATE": "__EMPTY_4",
    "REFERRAL": "__EMPTY_8",
    "REFERRAL2": "__EMPTY_9",
    "VISITOR": "__EMPTY_13",
    "TYFCB": "__EMPTY_15",
    "CEU": "__EMPTY_16",
    "RECOMMENDATION": "__EMPTY_17",
    "REPORT_DATE": "__EMPTY_9",
    "REPORT_DATE_NICE": "REPORT_DATE_NICE",

    "LAST_NAME_NICE": "LAST_NAME_NICE",
    "FIRST_NAME_NICE": "FIRST_NAME_NICE",
    "ATTENDANCE_NICE": "ATTENDANCE_NICE",
    "ABSENT_NICE": "ABSENT_NICE",
    "LATE_NICE": "LATE_NICE",
    "REFERRAL_NICE": "REFERRAL_NICE",
    "REFERRAL2_NICE": "REFERRAL2_NICE",
    "REFERRAL_FINAL": "REFERRAL_FINAL",
    "VISITOR_NICE": "VISITOR_NICE",
    "TYFCB_NICE": "TYFCB_NICE",
    "CEU_NICE": "CEU_NICE",
    "RECOMMENDATION_NICE": "RECOMMENDATION_NICE",
  };

  const EXCLUDE_DATA = ['ビジター', 'BNI', '合計'];

  let g_xlsxData = [];
  let g_members = {};

  $(".file-multiple").change(function (event) {
    reset();

    let files = event.target.files;
    if (files.length > 6) {
      showMessage(MSG_MAXIMUM_FILES);
    }

    for (let i = 0, len = files.length; i < len; i++) {
      readFile(files[i]);
    }

    // Load data in delay time
    setTimeout(function () {
      handleFileInput();
    }, 1000);
  })

  $(".export-score").click(function () {
    exportXLSX();
  });

  function handleFileInput() {
    console.log(g_xlsxData);
    // List column from input file:
    // Col 1: チャプター► サマリーPALMSレポート: "姓"    -- mapping --
    // Col 2: __EMPTY: "名"    -- mapping --
    // Col 3: __EMPTY_1: "出"    -- mapping --
    // Col 4: __EMPTY_2: "欠"    -- mapping --
    // Col 5: __EMPTY_4: "遅"    -- mapping --
    // Col 6: __EMPTY_5: "医"    -- mapping --
    // Col 7: __EMPTY_6: "代"    -- mapping --
    // Col 8: __EMPTY_8: "内与"    -- mapping --
    // Col 9: __EMPTY_9: "外与"    -- mapping --
    // Col 10: __EMPTY_10: "内受"    -- mapping --
    // Col 11: __EMPTY_12: "外受"    -- mapping --
    // Col 12: __EMPTY_13: "ビ"    -- mapping --
    // Col 13: __EMPTY_14: "1to1"    -- mapping --
    // Col 14: __EMPTY_15: "千円（千円未満切り捨て）"    -- mapping --
    // Col 15: __EMPTY_16: "CEU"    -- mapping --
    // Col 16: __EMPTY_17: "推薦状"    -- mapping --

    let fullName;
    let referral_final;
    let lastName; // Key: チャプター► サマリーPALMSレポート
    let firstName; // Key: __EMPTY
    let attendance; // Key: __EMPTY_1
    let absent; // Key: __EMPTY_2
    let late; // Key: __EMPTY_4
    let referral; // Key: __EMPTY_8
    let referral2; // Key: __EMPTY_9
    let visitor; // Key: __EMPTY_13
    let tyfcb; // Key: __EMPTY_15
    let ceu; // Key: __EMPTY_16
    let recommendation; // Key: __EMPTY_17
    let reportDate;

    let dataByDate = {};

    // Combine data from all files
    let data = [];
    for (let i = 0, len = g_xlsxData.length; i < len; i++) {
      let fileData = g_xlsxData[i];
      let _reportDate; // Key: __EMPTY_9

      // Get report date
      _reportDate = fileData[4];
      _reportDate = _reportDate[KEY.REPORT_DATE];
      // _reportDate = extractMonth(_reportDate);

      if (!dataByDate[_reportDate]) {
        dataByDate[_reportDate] = fileData;
      }

      // We don't need to calculate 3 item: 'ビジター', 'BNI', '合計'
      for (let j = 0, len2 = fileData.length - 3; j < len2; j++) {
        if (j < 7) continue;

        fileData[j][KEY.REPORT_DATE_NICE] = _reportDate;

        data.push(fileData[j]);
      }
    }

    console.log(data);

    // Calculate data of each member
    for (let i = 0, len = data.length; i < len; i++) {
      let row = data[i];

      lastName = row[Object.keys(row)[0]];
      firstName = row[KEY.FIRST_NAME];
      fullName = `${lastName} ${firstName}`;
      attendance = Number(row[KEY.ATTENDANCE]);
      absent = Number(row[KEY.ABSENT]);
      late = Number(row[KEY.LATE]);
      referral = Number(row[KEY.REFERRAL]);
      referral2 = Number(row[KEY.REFERRAL2]);
      referral_final = referral + referral2;
      visitor = Number(row[KEY.VISITOR]);
      tyfcb = Number(row[KEY.TYFCB]);
      ceu = Number(row[KEY.CEU]);
      recommendation = Number(row[KEY.RECOMMENDATION]);
      reportDate = row[KEY.REPORT_DATE_NICE];

      if (!g_members[fullName]) {
        g_members[fullName] = {};
      } else {
        attendance += g_members[fullName][KEY.ATTENDANCE_NICE];
        absent += g_members[fullName][KEY.ABSENT_NICE];
        late += g_members[fullName][KEY.LATE_NICE];
        referral += g_members[fullName][KEY.REFERRAL_NICE];
        referral2 += g_members[fullName][KEY.REFERRAL2_NICE];
        referral_final += g_members[fullName][KEY.REFERRAL_FINAL];
        visitor += g_members[fullName][KEY.VISITOR_NICE];
        tyfcb += g_members[fullName][KEY.TYFCB_NICE];
        ceu += g_members[fullName][KEY.CEU_NICE];
        recommendation += g_members[fullName][KEY.RECOMMENDATION_NICE];
      }

      g_members[fullName][KEY.LAST_NAME_NICE] = lastName;
      g_members[fullName][KEY.FIRST_NAME_NICE] = firstName;
      g_members[fullName][KEY.ATTENDANCE_NICE] = attendance;
      g_members[fullName][KEY.ABSENT_NICE] = absent;
      g_members[fullName][KEY.LATE_NICE] = late;
      g_members[fullName][KEY.REFERRAL_NICE] = referral;
      g_members[fullName][KEY.REFERRAL2_NICE] = referral2;
      g_members[fullName][KEY.REFERRAL_FINAL] = referral_final;
      g_members[fullName][KEY.VISITOR_NICE] = visitor;
      g_members[fullName][KEY.TYFCB_NICE] = tyfcb;
      g_members[fullName][KEY.CEU_NICE] = ceu;
      g_members[fullName][KEY.RECOMMENDATION_NICE] = recommendation;

      if (!g_members[fullName][KEY.REPORT_DATE_NICE]) {
        g_members[fullName][KEY.REPORT_DATE_NICE] = {};
      }

      if (!g_members[fullName][KEY.REPORT_DATE_NICE][reportDate]){
        g_members[fullName][KEY.REPORT_DATE_NICE][reportDate] = {};
      }

      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.ATTENDANCE_NICE] = Number(row[KEY.ATTENDANCE]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.ABSENT_NICE] = Number(row[KEY.ABSENT]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.LATE_NICE] = Number(row[KEY.LATE]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL_NICE] = Number(row[KEY.REFERRAL]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL2_NICE] = Number(row[KEY.REFERRAL2]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL_FINAL] = Number(row[KEY.REFERRAL]) + Number(row[KEY.REFERRAL2]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.VISITOR_NICE] = Number(row[KEY.VISITOR]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.TYFCB_NICE] = Number(row[KEY.TYFCB]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.CEU_NICE] = Number(row[KEY.CEU]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.RECOMMENDATION_NICE] = Number(row[KEY.RECOMMENDATION]);
    }

    console.log(g_members);
  }

  function updateExportButtonState(isValid) {
    isValid = isValid || false;
    if (isValid) {
      $(".export-score").removeAttr("disabled");
    } else {
      $(".export-score").attr("disabled");
    }
  }

  function showMessage(msg) {
    if (!msg || msg == '') {
      $(".message").hide();
      return;
    }

    $(".message__text").text(msg);
    $(".message").show();
  }

  function readFile(file) {
    const xl2json = new ExcelToJSON();
    xl2json.parseExcel(file);
  }

  function reset() {
    g_xlsxData = [];
    g_members = [];
  }

  function exportXLSX() {
    // Don't export if empty data
    // alert("export");
    ruleOfScore = new RuleOfScore(KEY);
    score = ruleOfScore.getScore(KEY.ABSENT, 1);
    rankColor = ruleOfScore.getRankColor(score);
    console.log(score, rankColor);

    const workbook = XLSX.utils.book_new();
    const worksheet_data = [
      ['', '※1　今月なにも活動しない場合の合計数値'],
      [
        '', 
        '', 
        '', 
        '参加 Attendance', 
        '欠席 Absent', '', '', '', '', '', '', '', '', 
        'LATE', '', '', '', '', '', '', '', '', 
        'INTERNAL REFERRAL', '', '', '', '', '', '', '', '', 
        'EXTERNAL REFFERAL', '', '', '', '', '', '', '', '', 
        'REFERRAL', '', 
        'VISITOR', '', '', '', '', '', '', '', '', 
        'TYFCB', '', '', '', '', '', '', '', '', 
        'CEU', '', '', '', '', '', '', '', '', 
        'RECOMMENDATION', '', '', '', '', '', '', '', '', 
      ],
      [
        '', 
        '', 
        '', 
        'Meetings 回数', 
        '12月', '1月', '2月', '3月', '4月', '5月', '', 'Presents', '※1', // absent
        '12月', '1月', '2月', '3月', '4月', '5月', '', '現在', '※1', // late
        '12月', '1月', '2月', '3月', '4月', '5月', '', '', '※1', // internal referral
        '12月', '1月', '2月', '3月', '4月', '5月', '', '', '※1', // external referral
        '現在', '※1', // referral
        '12月', '1月', '2月', '3月', '4月', '5月', '', '現在', '※1', // visitor
        '12月', '1月', '2月', '3月', '4月', '5月', '', '現在', '※1', // tyfcb
        '12月', '1月', '2月', '3月', '4月', '5月', '', '現在', '※1', // ceu
        '12月', '1月', '2月', '3月', '4月', '5月', '', '現在', '※1', // recommendation
      ],
      [
        '', 
        '姓 Last name', 
        '名 First name', 
        '', 
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
        '', '', 
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
        '', '', '', '', '', '', '', 'Total', '',
      ],
    ];

    for (const fullName in g_members) {
      const member = g_members[fullName];
      let row = [];

      row.push(""); // A
      row.push(member[KEY.LAST_NAME_NICE]); // B
      row.push(member[KEY.FIRST_NAME_NICE]); // C

      row.push(member[KEY.ATTENDANCE_NICE]); // D

      // absent
      row.push("0"); // E
      row.push("0"); // F
      row.push("0"); // G
      row.push("0"); // H
      row.push("0"); // I
      row.push("0"); // J
      row.push(""); // K
      row.push(member[KEY.ABSENT_NICE]); // L = SUM(E:K)
      row.push(""); // M = L - E

      // late
      row.push("0"); // N
      row.push("0"); // O
      row.push("0"); // P
      row.push("0"); // Q
      row.push("0"); // R
      row.push("0"); // S
      row.push(""); // T
      row.push(member[KEY.LATE_NICE]); // U = SUM(N:T)
      row.push(""); // V = U - N

      // internal referral
      row.push("0"); // W
      row.push("0"); // X
      row.push("0"); // Y
      row.push("0"); // Z
      row.push("0"); // AA
      row.push("0"); // AB
      row.push(""); // AC
      row.push(member[KEY.REFERRAL_NICE]); // AD = SUM(W:AC)
      row.push(""); // AE = AD - W

      // external referral
      row.push("0"); // AF
      row.push("0"); // AG
      row.push("0"); // AH
      row.push("0"); // AI
      row.push("0"); // AJ
      row.push("0"); // AK
      row.push(""); // AL
      row.push(member[KEY.REFERRAL2_NICE]); // AM = SUM(AF:AL)
      row.push(""); // AN = AM - AF

      // referral
      row.push("0"); // AO = AD + AM
      row.push("0"); // AP = AE5 + AN5

      // visitor
      row.push("0"); // AQ
      row.push("0"); // AR
      row.push("0"); // AS
      row.push("0"); // AT
      row.push("0"); // AU
      row.push("0"); // AV
      row.push(""); // AW
      row.push(member[KEY.VISITOR_NICE]); // AX = SUM(AQ:AW)
      row.push(""); // AY = AX - AQ

      // tyfb
      row.push("0"); // AZ
      row.push("0"); // BA
      row.push("0"); // BB
      row.push("0"); // BC
      row.push("0"); // BD
      row.push("0"); // BE
      row.push(""); // BF
      row.push(member[KEY.TYFCB_NICE]); // BG = SUM(AZ:BF)
      row.push(""); // BH = BG - AZ

      // CEU
      row.push("0"); // BI
      row.push("0"); // BJ
      row.push("0"); // BK
      row.push("0"); // BL
      row.push("0"); // BM
      row.push("0"); // BN
      row.push(""); // BO
      row.push(member[KEY.CEU_NICE]); // BP = SUM(BI:BO)
      row.push(""); // BQ = BP - BI

      // recommendation
      row.push("0"); // BR
      row.push("0"); // BS
      row.push("0"); // BT
      row.push("0"); // BU
      row.push("0"); // BV
      row.push("0"); // BW
      row.push(""); // BX
      row.push(member[KEY.RECOMMENDATION_NICE]); // BY = SUM(BR:BX)
      row.push(""); // BZ = BY - BR

      worksheet_data.push(row);
    }

    const worksheet = XLSX.utils.aoa_to_sheet(worksheet_data);

    workbook.SheetNames.push("Test");
    workbook.Sheets["Test"] = worksheet;

    exportExcelFile(workbook);
  }

  class RuleOfScore {
    constructor(key) {
      this.KEY = key;
    }
    getScore = function (key, count) {
      let score = 0;
      if (key == this.KEY.ABSENT) {
        switch (count) {
          case 2:
            score = 5;
            break
          case 1:
            score = 10;
            break
          case 0:
            score = 15;
            break
          case 3:
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.LATE) {
        switch (count) {
          case 1:
            score = 0;
            break
          case 0:
            score = 5;
            break
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.REFERRAL_FINAL) {
        switch (count) {
          case 0.75:
            score = 5;
            break
          case 1:
            score = 10;
            break
          case 1.5:
            score = 15;
            break
          case 2:
            score = 20;
            break
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.VISITOR) {
        switch (count) {
          case 2:
            score = 5;
            break
          case 4:
            score = 10;
            break
          case 6:
            score = 15;
            break
          case 12:
            score = 20;
            break
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.TYFCB) {
        switch (count) {
          case 200:
            score = 5;
            break
          case 500:
            score = 10;
            break
          case 1000:
            score = 15;
            break
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.CEU) {
        switch (count) {
          case 10:
            score = 5;
            break
          case 15:
            score = 10;
            break
          case 20:
            score = 15;
            break
          default:
            score = 0;
            break
        }
      } else if (key == this.KEY.RECOMMENDATION) {
        switch (count) {
          case 1:
            score = 5;
            break
          case 2:
            score = 10;
            break
          default:
            score = 0;
            break
        }
      } else {
        // Do nothing
      }

      return score;
    }

    getRankColor = function (score) {
      if (score >= 70 && score <= 95) {
        return "green"
      } else if (score >= 50 && score <= 65) {
        return "yellow"
      } else if (score >= 30 && score <= 45) {
        return "red"
      } else if (score >= 0 && score <= 25) {
        return "black"
      } else {
        // Do nothing
      }

    }
  }

  class ExcelToJSON {
    parseExcel = function (file) {
      var reader = new FileReader();

      reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        workbook.SheetNames.forEach(function (sheetName) {
          // Here is your object
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);

          g_xlsxData.push(JSON.parse(json_object))
        });
      };

      reader.onerror = function (ex) {
        console.log(ex);
      };

      reader.readAsBinaryString(file);
    };
  }
})

function exportExcelFile(workbook, fileName = "bookName") {
  return XLSX.writeFile(workbook, `${fileName}.xlsx`);
}

/**
 * Get month from date string
 * @param   {string}  dateStr  Date string with format YYYY年MM月DD日
 * @return  {string}           return month with a character of month
 */
function extractMonth(dateStr) {
  const regex = /年\d+月/gm;

  // Alternative syntax using RegExp constructor
  // const regex = new RegExp('年\d+月', 'gm')

  let m;
  let month;

  while ((m = regex.exec(dateStr)) !== null) {
    // This is necessary to avoid infinite loops with zero-width matches
    if (m.index === regex.lastIndex) {
        regex.lastIndex++;
    }
    
    // The result can be accessed through the `m`-variable.
    m.forEach((match, groupIndex) => {
        month = match;
    });
  }

  month = month.replace("年", "");

  return month;
}
