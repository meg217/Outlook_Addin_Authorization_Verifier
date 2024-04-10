
// this file contains the banner etraction logic from the body as
// well as parsing the banner and retrieving the categories


/**
 * function to extract banner from message body
 * parameter is the message body contents
 * returns the banner from the body
 * @param { String } body
 */
function getBannerFromBody(body) {
    const banner_regex =
      /^(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im;
  
    const banner = body.match(banner_regex);
    console.log(banner);
    if (banner) {
      console.log("banner found");
      return banner[0];
    } else {
      console.log("banner null");
      return null;
    }
  }
  
  /**
   * function to parse banner markings
   * parameter is the banner
   * returns an array of each category being array[1] is cat1 and on for 1, 4 and 7
   * @param { String } banner
   */
  function parseBannerMarkings(banner) {
    // const cat1_regex = "TOP[\s]*SECRET|TS|(TS)|SECRET|S|(S)|CONFIDENTIAL|C|(C)|UNCLASSIFIED|U|(U)";
    // const cat4_regex = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS";
    // const cat7_regex = "ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
    // const cat4_and_cat7 = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
    const cat1_regex = /TOP\s*SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U/gi;
    const cat4_regex = /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS/gi;
    const cat7_regex =
      /ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
    const cat4_and_cat7 =
      /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
  
    const Categories = banner.split("//");
    console.log(Categories);
    let Category_1 = Category(Categories[0], cat1_regex, 1);
    let Category_4 = null;
    let Category_7 = null;
    if (Categories[1]) {
      if (Categories[1].toUpperCase().match(cat7_regex)) {
        // If the second parse matches the regex for category 7, then we need to make category 4 null and run category 7
        console.log("second category matches category 7");
        Category_4 = null;
        Category_7 = Category(Categories[1], cat7_regex, 7);
      } else {
        // If the second parse doesnt match, run each category with its corresponding regex
        console.log("second category doesnt match category 7, running normal program");
        Category_4 = Category(Categories[1], cat4_regex, 4);
        Category_7 = Category(Categories[2], cat7_regex, 7);
      }
    } else {
      console.log("second category returned null");
    }
  
    const Together = [Category_1, Category_4, Category_7];
    checkDisseminations(Category_1, Category_7);
    return Together;
  }
  
  /**
   * returns the submarkings of the category. if there is one category, then it returns null
   * @param { string } category
   * @returns { array } || null
   */
  function getSubMarkings(category) {
    submarkings = category.split("/");
    if (submarkings.length <= 1) {
      console.log("There is only one submarking");
      return null;
    }
    console.log(submarkings);
    return submarkings;
  }
  
  /**
   * function that uses regex to match the input category string, if no match is found it returns null
   * @param { String } category
   * @param { String } regex
   * @param { int } categoryNum
   */
  function Category(category, regex, categoryNum) {
    if (!category) {
      console.log("Category " + categoryNum + " string returned null");
      return null;
    } else if (category.toUpperCase().match(regex)) {
      console.log("returning category YAY " + categoryNum);
      console.log(category.toUpperCase());
      return category.toUpperCase();
    }
    console.log("String did not match category " + categoryNum + "'s regex");
    return null;
  }

  /**
   * 
   * @param {String} classification 
   * @param {String} dissemination 
   */
  function checkDisseminations(classification, dissemination) {
    console.log("CLASSIFICATION: " + classification + "\n");
    console.log("DISSEM: " + dissemination + "\n");
  }