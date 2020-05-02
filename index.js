require('dotenv').config()
const XLSX = require('xlsx')
const axios = require('axios')
const filename = './Word_list.xlsx'
const columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
const max_rows = 67
const raw = XLSX.readFile(filename)
const content = raw.Sheets.Sheet1

function makeQuizLet(i){
  const term = content[columns[0] + i].v
  const pronunciation = content[columns[1] + i] ? content[columns[1] + i].v : ''
  const wordType = content[columns[2] + i] ? content[columns[2] + i].v : ''
  const definition = content[columns[3] + i].v
  const translation = content[columns[4] + i] ? content[columns[4] + i].v : ''

  const termCell = `${term}\n(${wordType})\n${pronunciation}`

  const vietnamVer = {termCell, difinitionCell: translation}
  const englishVer = {termCell, difinitionCell: definition}
  return {vietnamVer, englishVer}
}

function makeSubmitBody(rank, word, definition){
  const {QUIZLET_SET_ID} = process.env
  const data = [{
    setId: QUIZLET_SET_ID,
    word,
    rank,
    definition
  }]
  const now = new Date().getTime()
  const requestId = now + ":term:op-seq-0"
  return {data, requestId}
}

let index = 1
function sendRequest({termCell, difinitionCell}){
  index++
  const body = makeSubmitBody(index, termCell, difinitionCell)
  const {COOKIE, CS_TOKEN} = process.env
  return axios.post(
    'https://quizlet.com/webapi/3.2/terms/save?_method=PUT',
    body,
    {
      headers: {
        'Cookie': COOKIE,
        'CS-Token': CS_TOKEN,
        'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:75.0) Gecko/20100101 Firefox/75.0'
      },
    }
  )
}

(async function(){
  for (let i=5; i <= max_rows;i++){
    const {vietnamVer, englishVer} = makeQuizLet(i)
    await sendRequest(vietnamVer)
    // await sendRequest(englishVer)
  }
})();