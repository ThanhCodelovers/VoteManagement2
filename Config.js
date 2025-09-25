const listSheet = {
    setting: '設定',
    questions: '投票設定', 
    responseForm: '回答管理雛形',
    menu: 'メニュー',
    responseTotal: '回答一覧',
    listVote: '投票リスト',
    templateVoteSetting: '投票設定サンプル',
    verifyManager: '利用者管理',
    loginOTP: 'ログインOTP',
    colors: 'カラー設定',
}

const headerVerifyManager = {
  no: 'No',
  fullname: '利用者',
  email: 'Eメール',
  phoneNumber: '電話番号',
  status: '状態',
  active: '有効',
  unactive: '無効',
}

const headerLoginOTP = {
  no: 'No',
  fullname: '利用者',
  email: 'Eメール',
  otp: 'OTP',
  fromTime: 'From',
  toTime: 'To',
}

const listTitle = {
    no: '#',
    question: '質問',
    criterias: '基準',
    question_type: '質問種類', 
    question_hasNote: '追加回答', 
    question_labelTextNote: '追加質問', 
    question_answers: '答え', 
    question_required: '必須',
    description_for_answer: '質問の説明',
    description_for_note: '追加の説明',
    answer_created: '作成時間',
    full_url: 'ユーザ登録URL',
    sort_url: '短縮URL',
    lat: '緯度', 
    long: '経度',
    dimensionAllow: '許容距離（メーター）',
    class_title: '授業のタイトル',
    survey_title: '投票のタイトル',
    survey_description: '投票の紹介内容',
    survey_datetime: '実施日時',
    address: '住所',
    phone_number: '電話番号',
    email: 'Eメール',
    name: '名前',
    allowed_ip: 'IPアドレス制限',
    linkToResponse: '生徒からの回答はこちら',
    display_type: '表示モード',
    colors: '色',
    voteOrder: '投票基準の選択',
    max: '最大数',
    min: '最小数',
    voteMethod: '投票方法',
    voteThreshold: '投票閾値',
}

const question_type = {
    input: "入力",
    oneSelect: "1 つだけ選択",
    multiSelect: "複数選択",
    email: 'Eメール',
    date: '日付',
    textarea: '短い段落',
    score: '点数',
    target: '対象者選択',
    getQuestionType(value) {
      for (const [k, v] of Object.entries(this)) {
        if (v === value) return k
      }
      return null
    }
}

const display_type = {
    all: "１ページで全て表示",
    one: "１ページで１問ずつ表示",
    getDisplayType(value) {
      for (const [k, v] of Object.entries(this)) {
        if (v === value) return k
      }
      return null
    }
}

const statistics_type = {
  public : "表示",
  private: "非表示",
  getStatusType(value) {
    for (const [k, v] of Object.entries(this)) {
      if (v === value) return k
    }
    return null
  }
}
const status_type = {
  prepare : "未開始",
  processing: "公開中",
  completed:"終了済",
  getStatusType(value) {
    for (const [k, v] of Object.entries(this)) {
      if (v === value) return k
    }
    return null
  }
}

const master_sheet = {
  id: "コード",
  nameVote: "投票の名前",
  displayType: "表示モード",
  author: "担当者名",
  dateRun: "実施日",
  dateEnd: "終了日",
  status: "状態",
  statistics:"投票結果公共",
  description: "説明",
  url: "短縮URL",
  voteSetting: '投票設定',
	voteResponse: '回答一覧',
  numberQuestion: '質問数',
  numberVoted: '投票人数',
  informationRequired: '情報収集',
  usePassCode: 'パスコード入力',
  passcode: "パスコード",
  getAttributes() {
    return Object.keys(this).filter(key => typeof this[key] !== 'function');
  },
  getIndex(attribute) {
    const keys = Object.keys(this).filter(key => typeof this[key] !== 'function');
    const index = keys.indexOf(attribute);
    return index !== -1 ? index : null;
  },
  getLabel(attribute) {
    return this[attribute] !== undefined ? this[attribute] : null;
  },
  getLength() {
    return Object.keys(this).filter(key => typeof this[key] !== 'function').length;
  },
  getKeyByLabel(label) {
    for (const key in this) {
      if (this[key] === label) return key;
    }
    return null;
  },
  toObject(valueArray) {
    let keys = this.getAttributes();
    let result = {};
    if (keys.length !== valueArray.length) {
        throw new Error("The number of keys and values must match.");
    }
    keys.forEach((key, index) => {
        result[key] = valueArray[index];
    });

    return result;
  }
}

const headerSetting = {
	gas_url: 'ユーザ登録URL',
	quizListUrl: '短縮URL',
  OTPlifetime: 'OTPの有効期間',
	getDataByKey(key) {
		let res = null;
		try {
			let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.setting);
      let rows = s.getRange(5, 1, s.getLastRow() - 4, s.getMaxColumns()-2).getDisplayValues();
			for (let [k, v] of Object.entries(this)) {
				if (k !== key) continue

				let indexRow = rows.findIndex(row => {
					let idx = row.indexOf(v);
					if (idx === -1) return false
					return true
				});
	
				if (indexRow === -1) throw `${v} not found.`

				let indexValue = rows[indexRow].findIndex(x => x !== v && x !== '');
				if (indexValue === -1) throw `${v} is empty.`
				else {
					res = rows[indexRow][indexValue];
					break;
				}
			}
		} catch(error) {
			console.log(`get data setting by key got error: ${error}`);
			res = null
		} finally {
			return res
		}
	}
}