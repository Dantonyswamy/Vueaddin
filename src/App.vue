<template>
  <div id="app">
    <div id="content">
      <div id="content-header">
        <div class="padding"><h1>Welcome</h1></div>
      </div>
      <div id="content-main">
        <div class="padding">
          <p>
            Choose the button below to set the color of the selected range to
            green.
          </p>

          <h3>Try it out</h3>
          <v-btn @click="onSetColor">Set color</v-btn>
          <h3>Copy existing sheet to a new sheet</h3>
          <button @click="copytoNewsheet">Copy</button> {{ summary }}
          <h3>Copy existing sheet to a new sheet</h3>
          <button @click="createTable">Create Table</button>

          <div>{{ summary }}</div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import axios from "axios"

export default {
  name: "App",
  data() {
    return {
      summary: ""
    }
  },
  mounted() {
    axios
      .get("https://api.coindesk.com/v1/bpi/currentprice.json")
      .then(response => (this.summary = response))
  },
  methods: {
    onSetColor() {
      window.Excel.run(async context => {
        const range = context.workbook.getSelectedRange()
        range.format.fill.color = "green"
        await context.sync()
      })
    },
    copytoNewsheet() {
      window.Excel.run(async context => {
        let myWorkbook = context.workbook
        let sampleSheet = myWorkbook.worksheets.getActiveWorksheet()
        let copiedSheet = sampleSheet.copy("End")

        sampleSheet.load("name")
        copiedSheet.load("name")

        await context.sync()

        console.log(
          "'" + sampleSheet.name + "' was copied to '" + copiedSheet.name + "'"
        )
      })
    },
    createTable() {
      window.Excel.run(async context => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
        const expensesTable = currentWorksheet.tables.add(
          "A1:D1",
          true /*hasHeaders*/
        )
        expensesTable.name = "ExpensesTable"

        expensesTable.getHeaderRowRange().values = [
          ["Date", "Merchant", "Category", "Amount"]
        ]

        expensesTable.rows.add(null /*add at the end*/, [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        ])

        expensesTable.columns.getItemAt(3).getRange().numberFormat = [
          ["€#,##0.00"]
        ]
        expensesTable.getRange().format.autofitColumns()
        expensesTable.getRange().format.autofitRows()

        return context.sync()
      }).catch(function(error) {
        console.log("Error: " + error)
      })
    }
  }
}
</script>

<style>
#content-header {
  background: #2a8dd4;
  color: #fff;
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 80px;
  overflow: hidden;
}

#content-main {
  background: #fff;
  position: fixed;
  top: 80px;
  left: 0;
  right: 0;
  bottom: 0;
  overflow: auto;
}

.padding {
  padding: 15px;
}
</style>
