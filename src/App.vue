<template>
  <div id="app">
    <button @click="downloadPPT">生成自定义ppt</button>
    <p class="mb30" />
    <button  @click="genTablePPT">table生成ppt</button>
    <table id="tabAutoPaging">
      <thead>
      <tr>
        <th  style="width: 10%">Row</th>
        <th style="width:20%">Last Name</th>
        <th  style="width:10%">First Name</th>
        <th  style="width:70%">Description</th>
      </tr>
      </thead>
      <tbody>
        <tr>
          <td>1</td>
          <td>llll</td>
          <td>cccc</td>
          <td>desc</td>
        </tr>
        <tr>
          <td>2</td>
          <td>llll</td>
          <td>cccc</td>
          <td>desc</td>
        </tr>
      </tbody>
    </table>
    <p class="mb30" />
    <button  @click="genChartPPT">生成图表ppt</button>
  </div>
</template>

<script>
import pptxgen from "pptxgenjs";
export default {
  name: 'App',
  data(){
    return {
      pres: new pptxgen()
    }
  },
  methods: {
    downloadPPT(){
      let pres = this.pres;
      let textStyle =  { x: 1, y: 1, color: '363636', fill: { color:'F1F1F1' }, align: pres.AlignH.center }
      pres.layout = 'LAYOUT_WIDE'

      let slide1 = pres.addSlide();
      slide1.addText("我是第一页", textStyle);
      slide1.background = {fill: "#bd4029"}

      let slide2 = pres.addSlide();
      slide2.addImage({ path: require('./assets/img1.png'), x: 1, y: 2, w: "50%", h: "50%" })
      slide2.addText("我是第二页", textStyle);

      pres.writeFile("myPPT.pptx").then(fileName =>{
        console.log(`${fileName} created success`)
      })
    },
    genTablePPT(){
      let pres = this.pres;
      pres.tableToSlides("tabAutoPaging",{ master: "MASTER_SLIDE" });
      pres.writeFile("tableGen.pptx");
    },
    genChartPPT(){
      let pres = this.pres;
      let dataChartAreaLine = [
        {
          name: "Actual Sales",
          labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
          values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121],
        },
        {
          name: "Projected Sales",
          labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
          values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121],
        },
      ];

      let slide = pres.addSlide();
      slide.addChart(pres.ChartType.line, dataChartAreaLine, { x: 1, y: 1, w: 8, h: 4 });
      pres.writeFile("chartGen.pptx");
    }
  }
}
</script>

<style>
#app {
  width: 100%;
}
.mb30{
  margin-bottom: 30px;
}
table{
  border:1px solid #ccc;
  border-collapse: collapse;
}
tr,td,th{
  border:1px solid #ccc;
  padding:5px;
}
</style>
