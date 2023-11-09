const xlsx =require("xlsx"); 
const path = require('path');
const fs = require('fs'); 

const source_xlsx = "./storage/data.xlsx";

// get_question() 
// each question <question_block>
/*
{
  'ques': {
    'text': '',
    'img': ''
  }, 
  'opt': [], 'a, b, c, d'
  'ans': <opt_index> 2
} */

/*
  category = [
    'Road sign', 'general question'
  ]
*/

// each question category <category_block>
/*
  <cat_ind> = [questions] // [{question_block}, {question_block}, {question_block}]
/

/**
 study = [categories] // [{category_block}, {category_block}, {category_block}]
test_question = [categories] // [{category_block}, {category_block}, {category_block}]
*/

/* 
<state_index> = [study, test_question]
*/

class ExcelToJSON {
  #data_workbook;
  #settings_worksheet;
  #states_worksheet;
  #categories_list = [];
  #states_list_obj = {};
  #ans_mark = 'Yes';
  #op_json_dirname = 'json';
  //#op_json_dirpath = '';

  constructor() {
    let _source_xlsx_path = path.resolve(source_xlsx);
    this.#data_workbook = xlsx.readFile(_source_xlsx_path);
    this.#settings_worksheet = this.#data_workbook.Sheets['Settings'];
    this.#states_worksheet = this.#data_workbook.Sheets['States'];

    fs.mkdir(path.join(__dirname, this.#op_json_dirname), (err) => { 
      if (err) { 
        return console.error(err); 
      } 
      console.log('Directory created successfully!'); 
    });

    this.#fetch_categories();
    this.#fetch_states();
  }// end function

  #fetch_categories() {
    let _row_inc = 3;
    let _max_break_flag = 50;
    while(true) {
      let ind = 'B' + _row_inc++;
      let cell = this.#settings_worksheet[ind];
      
      if (typeof cell === 'undefined' || _max_break_flag-- <= 0) break;
      this.#categories_list.push(cell.v.toString().trim());

    }// end while  
    
    const question_categories = this.#op_json_dirname + '/question_categories.json';
    try {
      fs.writeFileSync(question_categories, JSON.stringify(this.#categories_list, null, 2), 'utf8');
      console.log('Data successfully saved to disk');
    } catch (error) {
      console.log('An error has occurred ', error);
    }

  }// end function

  #fetch_states() {
    let _row_inc = 3;
    let _max_break_flag = 50;
    while(true) {
      let state_title_ind = 'B' + _row_inc;
      let state_abbr_ind = 'C' + _row_inc;
      let state_flag_ind = 'D' + _row_inc;
      _row_inc++;

      let title_cell = this.#states_worksheet[state_title_ind];
      let abbr_cell = this.#states_worksheet[state_abbr_ind];
      let flag_cell = this.#states_worksheet[state_flag_ind];
      
      if ((typeof title_cell === 'undefined' && typeof abbr_cell === 'undefined') || _max_break_flag-- <= 0) break;

      let flag_img = '';
      if (typeof flag_cell !== 'undefined' && typeof flag_cell.f !== 'undefined') {
        let pre_Quote = flag_cell.f.toString().indexOf('"');
        let post_Quote = flag_cell.f.toString().indexOf('"', pre_Quote+1);
        flag_img = flag_cell.f.toString().substring(pre_Quote+1, post_Quote);
      }

      let state_tmp_info = {
        'title': title_cell.v.toString().trim(),
        'flag': flag_img
      }
      this.#states_list_obj[abbr_cell.v.toString().trim()] = state_tmp_info;

    }// end while   

    const states_file = this.#op_json_dirname + '/states.json';
    try {
      fs.writeFileSync(states_file, JSON.stringify(this.#states_list_obj, null, 2), 'utf8');
      console.log('Data successfully saved to disk');
    } catch (error) {
      console.log('An error has occurred ', error);
    }
    
  }// end function


  fetch_states_ques_ans() {
    // loop on <this.#states_list_obj>
    // console.log(this.#states_list_obj);
    this.#fetch_single_state_ques_ans('AB');

  }// end function


  #fetch_single_state_ques_ans(state_abbr) {
    let _state_qa_worksheet = this.#data_workbook.Sheets[state_abbr];
    let state_ques_ans_obj = {};

    if (typeof _state_qa_worksheet !== 'undefined') {
      let row_ind = 3;
      let _max_break_flag = 50;//2000;
      let qa_obj_list = [];

      while(true) {

        let [cat_index, cat_title] = this.#read_category(row_ind, _state_qa_worksheet);
        if (cat_title == '' || _max_break_flag-- <= 0) break;

        row_ind++;        
        [qa_obj_list, row_ind] = this.#read_question_answer_list(row_ind, _state_qa_worksheet);
        row_ind+=2;

        state_ques_ans_obj['cat-' + cat_index] = {
          'title': cat_title,
          'question_answer_list': qa_obj_list
        }

        // console.log(cat_index, cat_title);
        // console.log('QA: ');
        // console.log(row_ind, qa_obj_list);
        
      }// end while

      const states_qa_info = this.#op_json_dirname + '/'+ state_abbr +'-states_qa_info.json';
      try {
        fs.writeFileSync(states_qa_info, JSON.stringify(state_ques_ans_obj, null, 2), 'utf8');
        console.log('Data successfully saved to disk');
      } catch (error) {
        console.log('An error has occurred ', error);
      }
    }

    return state_ques_ans_obj;

    // console.log(_state_qa_worksheet);
  }// end function

  #read_category(ind, qa_sheet) {
    let cell_ind = 'A' + ind;
    let cat_title = typeof qa_sheet[cell_ind] !== 'undefined' ? qa_sheet[cell_ind].v : '';
    cat_title = cat_title.toString().trim();
    let cat_index = cat_title ? this.#categories_list.findIndex(ele => ele == cat_title) : -1;   

    return [cat_index, cat_title];
  }// end function

  #read_question_answer_list(ind, qa_sheet) {
    let qa_obj_list = [];
    let _max_break_flag = 5000;
    let qa_obj = {};

    while(true) {
      [qa_obj, ind] = this.#fetch_single_question(ind, qa_sheet);
      if (Object.keys(qa_obj).length == 0 || _max_break_flag-- <= 0) break;

      qa_obj_list.push(qa_obj);
      ind+=2;
    }// end while    

    return [qa_obj_list, ind];
  }// end function


  #fetch_single_question(ind, qa_sheet) {
    let qa_info = {};
    /*
    {
      'ques': {
        'text': '',
        'img': ''
      }, 
      'opt': [], 'a, b, c, d'
      'ans': <opt_index> 2
    } */
    // ind = 751;
    let cell_ind = 'C' + ind;
    let question_title = typeof qa_sheet[cell_ind] !== 'undefined' ? qa_sheet[cell_ind].v : '';
    let question_img = '';

    question_title = question_title.toString().trim();

    if(question_title) {
      // fetch question
      ind++;
      cell_ind = 'C' + ind;
      let img_cel = qa_sheet[cell_ind];
      if (typeof img_cel !== 'undefined' && typeof img_cel.f !== 'undefined') {
        let _pre_Quote = img_cel.f.toString().indexOf('"');
        let _post_Quote = img_cel.f.toString().indexOf('"', _pre_Quote+1);
        question_img = img_cel.f.toString().substring(_pre_Quote+1, _post_Quote);
        question_img = question_img.toString().trim();
      }
      qa_info['question'] = {
        'text': question_title,
        'img': question_img
      }

      // fetch option with answer            
      let max_ans_option = 20;
      let tmp_options = [];
      let ans_cell_index = '';
      let ans_cell_option_index_list = [];
      while(true) {
        ind++;
        cell_ind = 'D' + ind;
        ans_cell_index = 'I' + ind;
        let cell_info = qa_sheet[cell_ind];
        let ans_cell_info = qa_sheet[ans_cell_index];
        if (typeof cell_info === 'undefined' || max_ans_option -- <= 0) break;
        let cell_val = qa_sheet[cell_ind].v;

        if ((typeof cell_val === 'undefined' || cell_val == '') && typeof cell_info.f !== 'undefined') {
          let _pre_Quote = cell_info.f.toString().indexOf('"');
          let _post_Quote = cell_info.f.toString().indexOf('"', _pre_Quote+1);
          cell_val = cell_info.f.toString().substring(_pre_Quote+1, _post_Quote);
          cell_val = question_img.toString().trim();
        }
        cell_val = cell_val.toString().trim();
        if (cell_val) {
          tmp_options.push(cell_val);
          let ans_cell_val = typeof ans_cell_info !== 'undefined' ? ans_cell_info.v : '';
          if (ans_cell_val == this.#ans_mark) {
            ans_cell_option_index_list.push(tmp_options.length - 1);
          }
          
        }
        else break;
      }// end while
      qa_info['options'] = tmp_options;
      qa_info['answer_index'] = ans_cell_option_index_list;
    }

    // console.log(qa_info);
    // console.log('Ind->', ind);

    return [qa_info, ind];
  
  }// end function








  test_op() {
    // console.log(this.#categories_list);
    // console.log(this.#states_list);
  }// end function

}// end class

let excelToJSON = new ExcelToJSON();
excelToJSON.fetch_states_ques_ans();
excelToJSON.test_op();


