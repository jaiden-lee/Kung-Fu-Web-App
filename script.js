<script>
  let updateDataButton = document.getElementById("update-data");
  let displayDataButton = document.getElementById("display-data");
  let addDataButton = document.getElementById("add-data");
  let updateDataFrame = document.getElementById("update-data-frame");
  let displayDataFrame = document.getElementById("display-data-frame");
  let addDataFrame = document.getElementById("add-data-frame");
  let attendanceButton = document.getElementById("attendance-button");
  let excusedAbsenceButton = document.getElementById("excused-absence-button");
  let studentList = document.getElementById("student-list");
  let selectWeek = document.getElementById("select-week");
  let weekValues = document.getElementsByClassName("week-values");
  let duesSelect = document.getElementById("dues-type-select");
  let duesInput = document.getElementById("dues-input");
  let selectDuesMonth = document.getElementById("select-dues-month");
  let updateDuesButton = document.getElementById("update-dues-button");
  let remarksArea = document.getElementById("remarks-input-1");
  let updateRemarksButton = document.getElementById("update-remarks-button");
  let addStudentButton = document.getElementById("add-student-button");
  let addStudentInput = document.getElementById("add-student-input");
  let latestMonthLabel = document.getElementById("latest-month-label")
  let addMonthButton = document.getElementById("add-month-button");

  let studentDictionary = {};
  let monthList = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  let currentMonthIndex = -1;


  updateDataButton.onclick = function() {
    displayDataFrame.style.visibility = "hidden";
    updateDataFrame.style.visibility = "visible";
    addDataFrame.style.visibility = "hidden";
    google.script.run.withSuccessHandler(loadMonthsList).getMonthsList();
  }
  displayDataButton.onclick = function() {
    displayDataFrame.style.visibility = "visible";
    updateDataFrame.style.visibility = "hidden";
    addDataFrame.style.visibility = "hidden";
  }
  addDataButton.onclick = function() {
    updateDataFrame.style.visibility = "hidden";
    displayDataFrame.style.visibility = "hidden";
    addDataFrame.style.visibility = "visible";
    google.script.run.withSuccessHandler(updateLatestMonth).getLatestMonth();
  }

  function clearSelect (selectElement) {
    var i, L = selectElement.options.length - 1;
    for(i = L; i >= 0; i--) {
      selectElement.remove(i);
    }
  }

  function loadStudentList (nameList) {
    clearSelect(studentList);
    studentDictionary = nameList;
    let keys = Object.keys(nameList);
    for (let i=0; i<keys.length;i++) {
      let name = keys[i];
      let option = document.createElement("option");
      option.textContent= name;
      option.value = name;
      studentList.appendChild(option);   
    }
    updateAttendanceDisplay();
  }
  google.script.run.withSuccessHandler(loadStudentList).getNamesList();

  function getCurrentDay () {
    let today = new Date();
    let day = today.getDate();
    return day;
  }
  
  function setCurrentWeek () {
    let day = getCurrentDay();
    let week = Math.ceil(day/7);
    selectWeek.value = week;
  }
  setCurrentWeek();

  function updateAttendanceBox (parameters) {
    let weekNumber = parameters[0];
    let value = parameters[1];
    let weekId = "week-"+weekNumber;
    let textObject = document.getElementById(weekId).querySelector(".week-values");
    textObject.textContent = value;
  }

  function updateAttendanceDisplay () {
    let student = studentDictionary[studentList.value];
    let weekNumber = parseInt(selectWeek.value);
    if (student) {
      for (let i=0;i<weekValues.length;i++) {
        let week = weekValues[i].parentNode.id;
        let weekNumber = week.substring(week.length-1);
        google.script.run.withSuccessHandler(updateAttendanceBox).getWeekAttendance(student, weekNumber, currentMonthIndex);
      }
    }
  }

  function updateAttendanceSuccess () {
    let student = studentDictionary[studentList.value];
    let weekNumber = parseInt(selectWeek.value);
    if (student) {
      google.script.run.withSuccessHandler(updateAttendanceBox).getWeekAttendance(student, weekNumber, currentMonthIndex);
    }
  }

  attendanceButton.onclick = function () {
    let student = studentDictionary[studentList.value];
    let weekNumber = parseInt(selectWeek.value);
    if (student) {
      google.script.run.withSuccessHandler(updateAttendanceSuccess).updateAttendance(student, true, weekNumber, currentMonthIndex);
    }
  }

  excusedAbsenceButton.onclick = function () {
    let student = studentDictionary[studentList.value];
    let weekNumber = parseInt(selectWeek.value);

    if (student) {
      google.script.run.withSuccessHandler(updateAttendanceSuccess).updateAttendance(student, false, weekNumber, currentMonthIndex);
    }
  }

  function displayMonthlyDues (payment) {
    if (payment!=null && payment!="") {
      let paymentType = payment.substring(0,1);
      if (paymentType=="P") {
        duesSelect.value = "Paypal";
      } else {
        duesSelect.value = "Manual";
      }
      let paymentId = payment.substring(1).trim();
      duesInput.value = paymentId;
    } else {
      duesSelect.value = "Manual";
      duesInput.value = "";
    }
  }

  function loadMonthsList (months) {
    clearSelect(selectDuesMonth);
    let today = new Date();
    let currentMonth = monthList[today.getMonth()];
    let currentYear = today.getFullYear();
    for (let i=0;i<months.length;i++) {
      let text = months[i].text;
      let rowNumber = months[i].rowNumber;
      let option = document.createElement("option");
      option.textContent = text;
      option.value = rowNumber;
      selectDuesMonth.appendChild(option);
      if (text==currentMonth+" "+currentYear) {
        selectDuesMonth.value = parseInt(rowNumber);
        currentMonthIndex = parseInt(rowNumber);
      }
    }
  }
  google.script.run.withSuccessHandler(loadMonthsList).getMonthsList();

  updateDuesButton.onclick = function () {
    let student = studentDictionary[studentList.value];
    let month = selectDuesMonth.value;

    if (student) {
      google.script.run.updateMonthlyDues(student, duesSelect.value, duesInput.value, month);
    }
  }

  function displayRemarks (message) {
    remarksArea.value = message;
  }

  updateRemarksButton.onclick = function () {
    let student = studentDictionary[studentList.value];

    if (student) {
      google.script.run.updateRemarks(student, remarksArea.value);
    }
  }


  studentList.addEventListener("change", function() {
    updateAttendanceDisplay();

    let student = studentDictionary[studentList.value];

    if (student) {
      google.script.run.withSuccessHandler(displayMonthlyDues).getMonthlyDues(student, selectDuesMonth.value);  
      google.script.run.withSuccessHandler(displayRemarks).getRemarks(student); 
    }
  });

  selectDuesMonth.addEventListener("change", function() {
    let student = studentDictionary[studentList.value];

    if (student) {
      google.script.run.withSuccessHandler(displayMonthlyDues).getMonthlyDues(student, selectDuesMonth.value);
    }
  });

  function addStudentSuccess (success) {
    if (success) {
      alert("Student was successfully added");
    } else {
      alert("Student already exists in the database");
    }
  }

  addStudentButton.onclick = function() {
    let name = (addStudentInput.value).trim();
    if (name!="") {
      google.script.run.withSuccessHandler(addStudentSuccess).addStudent(name);
    } else {
      alert("Please enter a valid name");
    }
  }

  
  function updateLatestMonth(month) {
    latestMonthLabel.innerHTML = "<strong>Latest Month:</strong> "+month;
  }

  google.script.run.withSuccessHandler(updateLatestMonth).getLatestMonth();

   let canAddMonth = true;
  function addNewMonthSuccess (success) {
    canAddMonth = true;
    
    if (success==false) {
      alert("failure");
    } else {
      alert("success");
    }

    google.script.run.withSuccessHandler(updateLatestMonth).getLatestMonth();
  }

  addMonthButton.onclick = function() {
    if (canAddMonth) {
      canAddMonth = false;
      google.script.run.withSuccessHandler(addNewMonthSuccess).addNewMonth();
    }
  }

  google.script.run.checkToAddNewMonth();
</script>
