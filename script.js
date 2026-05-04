let certs = [];
let exps = [];
let skills = [];
let techSkills = [];
let softSkills = [];

/* HEADER */
document.getElementById("name").oninput = function(e){
  document.getElementById("p-name").innerText = e.target.value;
};

document.getElementById("contact").oninput = function(e){
  document.getElementById("p-contact").innerText = e.target.value;
};

document.getElementById("title").oninput = function(e){
  document.getElementById("p-title").innerText = e.target.value;
};

/* SUMMARY */
document.getElementById("summary").oninput = function(e){
  document.getElementById("p-summary").innerText = e.target.value;
};

/* EDUCATION */
document.getElementById("edu1").oninput = updateEdu;
document.getElementById("edu2").oninput = updateEdu;

function updateEdu(){
  document.getElementById("p-edu").innerHTML =
    document.getElementById("edu1").value + "<br>" +
    document.getElementById("edu2").value;
}

/* LANGUAGES */
document.getElementById("languages").oninput = function(e){
  const parts = e.target.value.split(",");
  document.getElementById("p-lang").innerHTML =
    parts.map(p => `<span>${p}</span>`).join("");
};

/* SKILLS */
function addSkill(){
  let i = document.createElement("input");
  i.placeholder = "Skill";
  i.oninput = renderSkills;
  document.getElementById("skillsBox").appendChild(i);
}

function addTech(){
  let i = document.createElement("input");
  i.placeholder = "Technical Skill";
  i.oninput = renderSkills;
  document.getElementById("techBox").appendChild(i);
}

function addSoft(){
  let i = document.createElement("input");
  i.placeholder = "Soft Skill";
  i.oninput = renderSkills;
  document.getElementById("softBox").appendChild(i);
}

function renderSkills(){
  skills = [...document.querySelectorAll("#skillsBox input")].map(i=>i.value);
  techSkills = [...document.querySelectorAll("#techBox input")].map(i=>i.value);
  softSkills = [...document.querySelectorAll("#softBox input")].map(i=>i.value);

  document.getElementById("p-skills").innerHTML = `
    ${skills.map(s=>`<div>${s}</div>`).join("")}
    ${techSkills.length ? "<br><b>Technical Skills</b>" : ""}
    ${techSkills.map(s=>`<div>${s}</div>`).join("")}
    ${softSkills.length ? "<br><b>Soft Skills</b>" : ""}
    ${softSkills.map(s=>`<div>${s}</div>`).join("")}
  `;
}

/* CERTIFICATIONS */
function addCert(){
  let i = document.createElement("input");
  i.placeholder = "Certification";
  i.oninput = updateCert;
  document.getElementById("certBox").appendChild(i);
}

function updateCert(){
  certs = [...document.querySelectorAll("#certBox input")].map(i=>i.value);
  document.getElementById("p-cert").innerHTML =
    certs.map(c=>`<div>• ${c}</div>`).join("");
}

/* EXPERIENCE */
function addExp(){
  let box = document.createElement("div");

  let t = document.createElement("input");
  t.placeholder = "Position";

  let c = document.createElement("input");
  c.placeholder = "Company | Duration";

  let d = document.createElement("textarea");
  d.placeholder = "Description";

  t.oninput = c.oninput = d.oninput = updateExp;

  box.append(t,c,d);
  document.getElementById("expBox").appendChild(box);
}

function updateExp(){
  exps = [];

  document.querySelectorAll("#expBox > div").forEach(b=>{
    let i = b.querySelectorAll("input,textarea");
    if(i.length >= 3){
      exps.push({
        t: i[0].value,
        c: i[1].value,
        d: i[2].value
      });
    }
  });

  document.getElementById("p-exp").innerHTML =
    exps.map(e=>`
      <div>
        <div class="exp-title">${e.t}</div>
        <div class="exp-company">${e.c}</div>
        <div class="exp-desc">${e.d}</div>
      </div>
    `).join("");
}

/* DOWNLOAD */
function downloadWord(){

const doc = new docx.Document({
sections:[{
children:[

new docx.Paragraph({
children:[new docx.TextRun({
text: document.getElementById("p-name").innerText,
bold:true,
size:48
})]
}),

new docx.Paragraph(document.getElementById("p-contact").innerText),
new docx.Paragraph(document.getElementById("p-title").innerText),

new docx.Paragraph(""),
new docx.Paragraph({children:[new docx.TextRun({text:"Profile Summary",bold:true})]}),
new docx.Paragraph(document.getElementById("p-summary").innerText),

new docx.Paragraph({children:[new docx.TextRun({text:"Work Experience",bold:true})]}),

...exps.flatMap(e=>[
new docx.Paragraph({children:[new docx.TextRun({text:e.t,bold:true})]}),
new docx.Paragraph(e.c),
new docx.Paragraph(e.d)
]),

new docx.Paragraph({children:[new docx.TextRun({text:"Skills",bold:true})]}),
new docx.Paragraph([...skills,...techSkills,...softSkills].join("\n")),

new docx.Paragraph({children:[new docx.TextRun({text:"Education",bold:true})]}),
new docx.Paragraph(document.getElementById("p-edu").innerText),

new docx.Paragraph({children:[new docx.TextRun({text:"Certifications",bold:true})]}),
new docx.Paragraph(certs.join("\n")),

new docx.Paragraph({children:[new docx.TextRun({text:"Languages",bold:true})]}),
new docx.Paragraph(document.getElementById("p-lang").innerText)

]
}]
});

docx.Packer.toBlob(doc).then(blob=>{
let a=document.createElement("a");
a.href=URL.createObjectURL(blob);
a.download="Eshal_CV.docx";
a.click();
});

}