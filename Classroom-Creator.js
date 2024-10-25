// string constants
const SHEET_NAMES = {
    MAIN: 'Sheet1',
    LOG: 'Sheet2'
  };
  
  const ACTION_TYPES = {
    CREATED: 'Created',
    UPDATED: 'Updated',
    NO_CHANGE: 'No changes needed',
    ERROR: 'Error'
  };
  
  const LOG_LEVELS = {
    INFO: 'INFO',
    WARNING: 'WARNING',
    ERROR: 'ERROR'
  };
  
  // main function to create or update classrooms
  function manageClassrooms() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = {
      main: spreadsheet.getSheetByName(SHEET_NAMES.MAIN),
      log: spreadsheet.getSheetByName(SHEET_NAMES.LOG)
    };
    
    if (!sheets.main || !sheets.log) {
      log(LOG_LEVELS.ERROR, 'main sheet or log sheet not found. please check the sheet names.');
      return;
    }
    
    ensureLogSheetInitialized(sheets.log);
    
    // ensure course id and last updated columns exist
    const updatedData = ensureColumnsExist(sheets.main);
    let data = updatedData ? updatedData : sheets.main.getDataRange().getValues();
    let headers = data[0];
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    let columnIndices = getColumnIndices(headers);
    
    // if columns were added, refresh the data and indices
    if (updatedData) {
      data = sheets.main.getDataRange().getValues();
      headers = data[0];
      columnIndices = getColumnIndices(headers);
    }
    
    const existingCourses = getCachedCourses();
    
    const updates = { main: [], log: [] };
    
    // process each row, skipping the header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // skip entirely empty rows
      if (row.every(cell => cell === '')) {
        log(LOG_LEVELS.INFO, `skipping empty row ${i + 1}`);
        continue;
      }
  
      const courseId = columnIndices.courseId !== -1 ? row[columnIndices.courseId] : null;
      const courseName = row[columnIndices.name];
      
      // skip processing for homeroom and intervention courses
      if (courseName && (courseName.includes('Homeroom') || courseName.includes('Intervention'))) {
        log(LOG_LEVELS.INFO, `skipping ${courseName} course`);
        continue;
      }
      
      log(LOG_LEVELS.INFO, `processing row ${i + 1}: ${courseId ? 'updating existing course' : 'creating new course'}`);
      
      try {
        const courseDetails = extractCourseDetails(row, columnIndices, currentUserEmail);
        const existingCourse = findExistingCourse(existingCourses, courseDetails);
        
        const { updatedCourse, action, updatedFields } = createOrUpdateCourse(existingCourse, courseDetails, currentUserEmail);
        
        if (updatedCourse && updatedCourse.id) {
          if (columnIndices.courseId !== -1 && !row[columnIndices.courseId]) {
            updates.main.push({ row: i + 1, col: columnIndices.courseId + 1, value: updatedCourse.id });
          }
          if (columnIndices.lastUpdated !== -1) {
            updates.main.push({ row: i + 1, col: columnIndices.lastUpdated + 1, value: new Date() });
          }
          
          const teacherResult = setTeacherAsHead(updatedCourse.id, courseDetails.teacherEmail, currentUserEmail);
          
          const studentResult = addStudentsToCourse(updatedCourse.id, courseDetails.students);
          
          let notes = `${action}. ${teacherResult.message} ${studentResult.message}`;
          if (updatedFields.length > 0) {
            notes += ` updated fields: ${updatedFields.join(', ')}.`;
          }
          updates.log.push([new Date(), courseDetails.name, `${action}`, updatedCourse.id, notes]);
        } else {
          log(LOG_LEVELS.ERROR, `invalid course id for ${courseDetails.name}`);
          updates.log.push([new Date(), courseDetails.name, ACTION_TYPES.ERROR, '', `invalid course id for ${courseDetails.name}`]);
        }
      } catch (error) {
        log(LOG_LEVELS.ERROR, `error in row ${i + 1}: ${error.message}`);
        handleError(error, row[columnIndices.name], updates.log, i + 1);
      }
    }
    
    // batch update main sheet if there are updates
    if (updates.main.length > 0) {
      try {
        updates.main.forEach(update => {
          sheets.main.getRange(update.row, update.col).setValue(update.value);
        });
        log(LOG_LEVELS.INFO, `updated ${updates.main.length} cells in main sheet`);
      } catch (error) {
        log(LOG_LEVELS.ERROR, `failed to update main sheet: ${error && error.message ? error.message : 'unknown error'}`);
      }
    }
    
    if (updates.log.length > 0) {
      try {
        sheets.log.getRange(sheets.log.getLastRow() + 1, 1, updates.log.length, updates.log[0].length).setValues(updates.log);
        log(LOG_LEVELS.INFO, `added ${updates.log.length} entries to log sheet`);
      } catch (error) {
        log(LOG_LEVELS.ERROR, `failed to update log sheet: ${error && error.message ? error.message : 'unknown error'}`);
      }
    }
  }
  
  // get column indices for relevant columns
  function getColumnIndices(headers) {
    return {
      name: headers.indexOf('Course/Class Name'),
      courseLeads: headers.indexOf('Course Leads this academic year'),
      subject: headers.indexOf('Subject'),
      yearGroup: headers.indexOf('Year Group'),
      courseId: headers.indexOf('Course ID'),
      lastUpdated: headers.indexOf('Last Updated'),
      students: headers.indexOf('Students')
    };
  }
  
  // extract course details from a row
  function extractCourseDetails(row, indices, currentUserEmail) {
    const name = row[indices.name];
    const teacherEmail = row[indices.courseLeads] || null;
    const subject = row[indices.subject];
    const yearGroup = row[indices.yearGroup];
    const studentEmails = indices.students !== -1 ? row[indices.students] : '';
    const extractedStudents = studentEmails.split(',').map(email => email.trim()).filter(email => email);
    
    // log the raw values for debugging
    log(LOG_LEVELS.INFO, `raw course data: name: "${name}", subject: "${subject}", teacher: "${teacherEmail}"`);
    log(LOG_LEVELS.INFO, `extracted ${extractedStudents.length} student emails for course: ${name}`);
    
    if (!name || name.trim() === '') {
      throw new Error(`course name is empty or undefined. raw value: "${name}"`);
    }
    
    return {
      name: name.trim(),
      teacherEmail: teacherEmail,
      description: `Subject: ${subject || 'N/A'}\nYear Group: ${yearGroup || 'N/A'}`,
      courseId: indices.courseId !== -1 ? row[indices.courseId] : null,
      students: extractedStudents
    };
  }
  
  // find an existing course by id or name
  function findExistingCourse(existingCourses, courseDetails) {
    return existingCourses.find(course => 
      course.id === courseDetails.courseId || 
      course.name === courseDetails.name
    );
  }
  
  // create a new course or update an existing one
  function createOrUpdateCourse(existingCourse, courseDetails, currentUserEmail) {
    const course = {
      name: courseDetails.name,
      descriptionHeading: courseDetails.name,
      description: courseDetails.description,
      courseState: 'ACTIVE'
    };
    
    let updatedCourse, action, updatedFields = [];
    
    if (existingCourse) {
      log(LOG_LEVELS.INFO, `existing course found: ${courseDetails.name}`);
      
      // determine which fields need updating
      const fieldsToUpdate = ['name', 'descriptionHeading', 'description']
        .filter(field => existingCourse[field] !== course[field]);
      
      // only update owner if it's explicitly provided and different
      if (courseDetails.teacherEmail && existingCourse.ownerId !== courseDetails.teacherEmail) {
        course.ownerId = courseDetails.teacherEmail;
        fieldsToUpdate.push('ownerId');
      }
      
      if (fieldsToUpdate.length > 0) {
        const updateMask = fieldsToUpdate.join(',');
        updatedCourse = Classroom.Courses.patch(course, existingCourse.id, {updateMask: updateMask});
        action = ACTION_TYPES.UPDATED;
        updatedFields = fieldsToUpdate;
        log(LOG_LEVELS.INFO, `course details updated for: ${courseDetails.name}. updated fields: ${updateMask}`);
      } else {
        updatedCourse = existingCourse;
        action = ACTION_TYPES.NO_CHANGE;
        log(LOG_LEVELS.INFO, `no updates needed for: ${courseDetails.name}`);
      }
    } else {
      log(LOG_LEVELS.INFO, `no existing course found. creating new course: ${courseDetails.name}`);
      course.ownerId = courseDetails.teacherEmail || currentUserEmail;
      updatedCourse = Classroom.Courses.create(course);
      action = ACTION_TYPES.CREATED;
      updatedFields = ['name', 'descriptionHeading', 'description', 'ownerId'];
    }
    
    return { updatedCourse, action, updatedFields };
  }
  
  // set teacher as head of the class only if there's a course lead
  function setTeacherAsHead(courseId, teacherEmail, currentUserEmail) {
    if (!teacherEmail) {
      return { success: true, message: 'no course lead specified, current user remains as owner.' };
    }
  
    try {
      Classroom.Courses.Teachers.create({userId: teacherEmail}, courseId);
      log(LOG_LEVELS.INFO, `set ${teacherEmail} as head teacher for course: ${courseId}`);
      return { success: true, message: `${teacherEmail} set as head teacher.` };
    } catch (error) {
      log(LOG_LEVELS.WARNING, `failed to set head teacher for course: ${courseId}. error: ${error && error.message ? error.message : 'unknown error'}`);
      return { success: false, message: `failed to set head teacher: ${error && error.message ? error.message : 'unknown error'}` };
    }
  }
  
  // add students to a course
  function addStudentsToCourse(courseId, studentEmails) {
    let added = 0, failed = 0;
    log(LOG_LEVELS.INFO, `attempting to add ${studentEmails.length} students to course ${courseId}`);
    
    for (const email of studentEmails) {
      try {
        Classroom.Courses.Students.create({userId: email}, courseId);
        added++;
        log(LOG_LEVELS.INFO, `successfully added student: ${email} to course: ${courseId}`);
      } catch (error) {
        failed++;
        log(LOG_LEVELS.WARNING, `failed to add student: ${email} to course: ${courseId}. error: ${error.message}`);
      }
    }
    
    const result = {
      success: added > 0,
      message: `added ${added} students. failed to add ${failed} students.`
    };
    log(LOG_LEVELS.INFO, `student addition result: ${result.message}`);
    return result;
  }
  
  // get cached courses or fetch them if cache is empty
  function getCachedCourses() {
    log(LOG_LEVELS.INFO, 'fetching courses from api');
    const courses = Classroom.Courses.list({teacherId: 'me'}).courses || [];
    return courses;
  }
  
  // ensure the log sheet is initialized with headers
  function ensureLogSheetInitialized(sheet) {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Course Name', 'Action', 'Course ID', 'Notes']);
    }
  }
  
  // ensure necessary columns exist in the main sheet
  function ensureColumnsExist(sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let updated = false;
  
    if (!headers.includes('Course ID')) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Course ID');
      updated = true;
      log(LOG_LEVELS.INFO, 'created course id column');
    }
  
    if (!headers.includes('Last Updated')) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Last Updated');
      updated = true;
      log(LOG_LEVELS.INFO, 'created last updated column');
    }
  
    if (updated) {
      // if columns were added, return the updated data
      return sheet.getDataRange().getValues();
    } else {
      return null;
    }
  }
  
  // log a message with a specified level
  function log(level, message) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] [${level}] ${message}`);
  }
  
  // handle errors that occur during processing
  function handleError(error, courseName, logs, rowIndex) {
    const errorMessage = `failed to process course: ${courseName} (row ${rowIndex}). error: ${error && error.message ? error.message : 'unknown error'}`;
    log(LOG_LEVELS.ERROR, errorMessage);
    logs.push([new Date(), courseName, ACTION_TYPES.ERROR, '', errorMessage]);
  }