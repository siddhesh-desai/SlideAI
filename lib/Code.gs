function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

var generatedUrls = [];
var list = [];

//Enter your details here
PRESENTATION_ID = "ENTER_YOUR_PRESENTATION_ID";
OPENAI_KEY = "Bearer ENTER_YOUR_OPENAI_KEY";
BING_API_KEY = "ENTER_YOUR_BING_SEARCH_API_KEY";



//Creating a new presentation
function createPresentation() {
  try {
    const presentation = Slides.Presentations.create({ title: "SlideAI" });
    console.log("Created presentation with ID: " + presentation.presentationId);
    return presentation.presentationId;
  } catch (e) {
    // TODO (developer) - Handle exception
    console.log("Failed with error %s", e.message);
  }
}

//Delete Slides
function deleteAllSlides(presentationId) {
  var presentation = SlidesApp.openById(presentationId);
  var slides = presentation.getSlides();

  // Iterate over each slide and delete it
  for (var i = slides.length - 1; i >= 0; i--) {
    var slide = slides[i];
    slide.remove();
    // presentation.removeSlide(slide);
  }
}

//Generate Link from Slides ID
function generateSlidesLink(presentationId) {
  return "https://docs.google.com/presentation/d/" + presentationId + "/edit";
}

//Main Function
function runn(
  project_name,
  project_description,
  color,
  author_names,
  institute_logo,
  occasion
) {

  //deleteAllSlides(PRESENTATION_ID);
  // const PRESENTATION_ID = createPresentation();
  // console.log(generateSlidesLink(PRESENTATION_ID));
  //Project Details
  // const project_name = "SlideAI";
  // const project_description = "SlideAI is an automated PowerPoint generator that streamlines the process of creating engaging presentations. By leveraging artificial intelligence, SlideAI takes care of content creation, slide design, and formatting, saving valuable time and effort. Simply input your key points, and SlideAI intelligently generates visually appealing slides with appropriate layouts, graphics, and text. It understands the context and presents information in a concise and captivating manner. With SlideAI, professionals can focus on delivering impactful presentations while eliminating the hassle of manual slide creation.";
  // const color = '#FF0000';
  // const author_names = 'XYZ';
  // const institute_logo = "https://img.collegepravesh.com/2017/05/VIT-Pune-Logo.png";
  // const occasion = 'EDI Review';

  // addTitleSlide(PRESENTATION_ID, project_name, color, occasion, 'Authors: ' + author_names);
  var slideIDDD = createSlide(PRESENTATION_ID, 0);
  addIntro(PRESENTATION_ID, project_name, project_description);
  addPS(PRESENTATION_ID, project_name, project_description);
  addMotivation(PRESENTATION_ID, project_name, project_description);
  addSolution(PRESENTATION_ID, project_name, project_description);
  addUnique(PRESENTATION_ID, project_name, project_description);
  addMethodology(PRESENTATION_ID, project_name, project_description);
  addTechnology(PRESENTATION_ID, project_name, project_description);
  addBenefits(PRESENTATION_ID, project_name, project_description);
  addLimitations(PRESENTATION_ID, project_name, project_description);
  addConclusion(PRESENTATION_ID, project_name, project_description);
  var slideIDD = createSlide(PRESENTATION_ID, 11);
  // addThankyouSlide(PRESENTATION_ID, color, author_names);

  addImagestoAllSlides(
    PRESENTATION_ID,
    project_description,
    color,
    institute_logo,
    project_name,
    occasion,
    author_names
  );
}

function generateDownloadLink(presentationId) {
  return (
    "https://docs.google.com/presentation/d/" + presentationId + "/export/pptx"
  );
}

function addTitleSlide(
  PRESENTATION_ID = "1ROXSzo28oiI8c6SSvsa1GXGruDkps9KpqlnoMqPLcQo",
  project_name = "SlideAI",
  color = "#FF0000",
  occasion = "EDI End Semester Examination",
  authors = "Authors: Siddhesh Desai, Ashish Fargade"
) {
  var slideID = createSlide(PRESENTATION_ID, 0);
  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];

  return;
}

function addThankyouSlide(PRESENTATION_ID, color, authors) {
  var slideIDD = createSlide(PRESENTATION_ID, 11);
  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[11];

  return;
}

function addImagestoAllSlides(
  PRESENTATION_ID = "1ROXSzo28oiI8c6SSvsa1GXGruDkps9KpqlnoMqPLcQo",
  project_description = "SlideAI is an automated PowerPoint generator that streamlines the process of creating engaging presentations. By leveraging artificial intelligence, SlideAI takes care of content creation, slide design, and formatting, saving valuable time and effort. Simply input your key points, and SlideAI intelligently generates visually appealing slides with appropriate layouts, graphics, and text. It understands the context and presents information in a concise and captivating manner. With SlideAI, professionals can focus on delivering impactful presentations while eliminating the hassle of manual slide creation.",
  color = "#FF0000",
  image_url = "https://img.collegepravesh.com/2017/05/VIT-Pune-Logo.png",
  project_name = "SlideAI",
  occasion = "Bla",
  authors = "sjbkas"
) {
  var presentation = SlidesApp.openById(PRESENTATION_ID);
  generateKeyWords(project_description);
  for (var i = 0; i < 13; i++) {
    var slide = presentation.getSlides()[i];
    if (i == 0) {
      addRectangleToSlide(slide, 0, 0, 750, 300, color);
      addTextToSlide(
        slide,
        50,
        90,
        650,
        80,
        project_name,
        50,
        "League Spartan",
        "#FFFFFF"
      );
      addTextToSlide(
        slide,
        50,
        160,
        650,
        60,
        occasion,
        30,
        "League Spartan",
        "#FFFFFF"
      );
      addTextToSlide(
        slide,
        50,
        330,
        650,
        60,
        authors,
        20,
        "League Spartan",
        "#000000"
      );
      addLogo(slide, image_url);
    } else if (i == 11) {
      addRectangleToSlide(slide, 0, 0, 750, 300, color);
      addTextToSlide(
        slide,
        50,
        90,
        650,
        80,
        "Thank you!",
        50,
        "League Spartan",
        "#FFFFFF"
      );
      addTextToSlide(
        slide,
        50,
        160,
        650,
        60,
        "Feel free to ask queries.",
        30,
        "League Spartan",
        "#FFFFFF"
      );
      addTextToSlide(
        slide,
        50,
        330,
        650,
        60,
        authors,
        20,
        "League Spartan",
        "#000000"
      );
      addLogo(slide, image_url);
    } else {
      addImage(slide, list[i % 10]);
      addRectangleToSlide(slide, 0, 0, 800, 10, color);
    }
  }
  return;
}

// Adding Introduction
function addIntro(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 1);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Write a curated introduction of 80 words for a presentation for the project named " +
    project_name +
    " which works as follows - " +
    project_description;
  addHeaderText(PRESENTATION_ID, slideID, "Introduction");
  content = truncateString(getInfo(prompt), 500);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Motivation
function addMotivation(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 3);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];

  // if (slide) {
  prompt =
    "Understand and write the motivation of making the project less than 80 words for a presentation for the project named " +
    project_name +
    " which works as follows - " +
    project_description;
  addHeaderText(PRESENTATION_ID, slideID, "Motivation");
  content = getInfo(prompt);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

//Add Problem Statement
function addPS(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 2);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];

  // if (slide) {
  prompt =
    "Understand and write the problem statements in bullets for a presentation for the project named " +
    project_name +
    " which works as follows - " +
    project_description;
  addHeaderText(PRESENTATION_ID, slideID, "Problem Statement");
  content = getInfo(prompt);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Our Solution
function addSolution(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 4);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the content which should be of 70 words for 'Our Solution' Slide of presentation for the project named" +
    project_name +
    " which works as follows - " +
    project_description;
  addHeaderText(PRESENTATION_ID, slideID, "Our Solution");
  content = getInfo(prompt);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding What's Unique?
function addUnique(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 5);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the unique features in bullet points for the project named" +
    project_name +
    " which works as follows - " +
    project_description +
    " It should be less than 80 Words";
  addHeaderText(PRESENTATION_ID, slideID, "What's Unique?");
  content = addSpaceAfterFullStop(getInfo(prompt));
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Methodology
function addMethodology(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 6);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the methodology in less than 80 words for the presentation of the project named" +
    project_name +
    " which works as follows - " +
    project_description;
  addHeaderText(PRESENTATION_ID, slideID, "Methodology");
  content = getInfo(prompt);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Technology Used
function addTechnology(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 7);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the technologies and softwares used in bullet points for the presentation of the project named" +
    project_name +
    " which works as follows - " +
    project_description +
    " It should be less than 80 Words";
  addHeaderText(PRESENTATION_ID, slideID, "Technology Used");
  content = addSpaceAfterFullStop(getInfo(prompt));
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Benefits
function addBenefits(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 8);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the benefits in bullet points for the presentation of the project named" +
    project_name +
    " which works as follows - " +
    project_description +
    " It should be less than 80 Words";
  addHeaderText(PRESENTATION_ID, slideID, "Benefits");
  content = addSpaceAfterFullStop(getInfo(prompt));
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Limitations
function addLimitations(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 9);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the limitations in bullet points for the presentation of the project named" +
    project_name +
    " which works as follows - " +
    project_description +
    " It should be less than 80 Words";
  addHeaderText(PRESENTATION_ID, slideID, "Limitations");
  content = addSpaceAfterFullStop(getInfo(prompt));
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

// Adding Conclusion
function addConclusion(PRESENTATION_ID, project_name, project_description) {
  var slideID = createSlide(PRESENTATION_ID, 10);

  // var presentation = SlidesApp.openById(PRESENTATION_ID);
  // var slide = presentation.getSlides()[0];
  // if (slide) {
  prompt =
    "Understand and write the conclusion for the presentation of the project named" +
    project_name +
    " which works as follows - " +
    project_description +
    " It should be less than 80 Words";
  addHeaderText(PRESENTATION_ID, slideID, "Conclusion");
  content = getInfo(prompt);
  addContentText(PRESENTATION_ID, slideID, content);
  // addImage(slide, content);
  // Utilities.sleep(5000);
  // } else {
  // console.log("Slide not found for ID: " + slideID);
  // }
  return;
}

//Add Header Text
function addHeaderText(presentationId, pageId, content) {
  //Get Unique ID
  const pageElementId = Utilities.getUuid();

  const requests = [
    {
      createShape: {
        objectId: pageElementId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: pageId,
          size: {
            width: {
              magnitude: 425,
              unit: "PT",
            },
            height: {
              magnitude: 60,
              unit: "PT",
            },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 36,
            translateY: 36,
            unit: "PT",
          },
        },
      },
    },
    {
      insertText: {
        objectId: pageElementId,
        text: content,
        insertionIndex: 0,
      },
    },
    {
      updateTextStyle: {
        objectId: pageElementId,
        fields: "foregroundColor,bold,italic,fontFamily,fontSize,underline",
        style: {
          foregroundColor: {
            opaqueColor: {
              themeColor: "ACCENT2",
            },
          },
          bold: true,
          italic: false,
          underline: false,
          fontFamily: "League Spartan",
          fontSize: {
            magnitude: 30,
            unit: "PT",
          },
        },
        textRange: {
          type: "ALL",
        },
      },
    },
  ];

  //Batch Update
  try {
    const response = Slides.Presentations.batchUpdate(
      {
        requests: requests,
      },
      presentationId
    );
    console.log(
      "Created Textbox with ID: " + response.replies[0].createShape.objectId
    );
    return response;
  } catch (e) {
    console.log("Failed with error %s", e.message);
  }
}

//Get Information from API
function getInfo(prompt) {
  var response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/engines/text-davinci-002/completions",
    {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        Authorization:
          OPENAI_KEY,
      },
      payload: JSON.stringify({
        prompt: prompt,
        max_tokens: 1024,
        temperature: 0,
        n: 1,
      }),
    }
  );

  answer = JSON.parse(response).choices[0].text.trim();
  console.log(answer);
  return answer;
}

//Add Content Text
function addContentText(presentationId, pageId, content) {
  //Get Unique ID
  const pageElementId = Utilities.getUuid();

  //Get Content Here
  // const content = 'sit amet justo donec enim diam vulputate ut pharetra sit amet aliquam id diam maecenas ultricies mi eget mauris pharetra et ultrices neque ornare aenean euismod elementum nisi quis eleifend quam adipiscing vitae proin sagittis nisl rhoncus mattis rhoncus urna neque viverra justo nec ultrices dui sapien eget mi proin sed libero enim sed faucibus turpis in eu mi bibendum neque egestas congue quisque egestas diam in arcu cursus euismod quis viverra nibh cras pulvinar mattis nunc sed blandit libero volutpat sed cras ornare arcu dui vivamus arcu felis bibendum ut tristique et egestas quis ipsum suspendisse ultrices gravida dictum';

  const requests = [
    {
      createShape: {
        objectId: pageElementId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: pageId,
          size: {
            width: {
              magnitude: 425,
              unit: "PT",
            },
            height: {
              magnitude: 300,
              unit: "PT",
            },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 36,
            translateY: 95,
            unit: "PT",
          },
        },
      },
    },
    {
      insertText: {
        objectId: pageElementId,
        text: content,
        insertionIndex: 0,
      },
    },
    {
      updateTextStyle: {
        objectId: pageElementId,
        fields: "foregroundColor,bold,italic,fontFamily,fontSize,underline",
        style: {
          foregroundColor: {
            opaqueColor: {
              themeColor: "ACCENT2",
            },
          },
          bold: false,
          italic: false,
          underline: false,
          fontFamily: "League Spartan",
          fontSize: {
            magnitude: 17,
            unit: "PT",
          },
        },
        textRange: {
          type: "ALL",
        },
      },
    },
  ];

  //Batch Update
  try {
    const response = Slides.Presentations.batchUpdate(
      {
        requests: requests,
      },
      presentationId
    );
    console.log(
      "Created Textbox with ID: " + response.replies[0].createShape.objectId
    );
    return response;
  } catch (e) {
    console.log("Failed with error %s", e.message);
  }
}

//Adding image on slide
function addImage(slide, image_text) {
  var image_url = fetchImageUrl(image_text);
  if (image_url != null) {
    console.log("Inserted Here");
    console.log(slide);
    console.log(image_url);
    try {
      var image = slide.insertImage(image_url.trim());
      image.setWidth(245);
      image.setHeight(405);
      image.setLeft(475);
      image.setTop(0);
      return;
    } catch (e) {
      console.log("Image not found!");
    }
  } else {
    console.log("Image URL is null");
  }
}

//Adding logo on slide
function addLogo(slide, image_url) {
  if (image_url != null) {
    console.log("Inserted Here");
    console.log(slide);
    console.log(image_url);
    var image = slide.insertImage(image_url.trim());
    image.setWidth(100);
    image.setHeight(100);
    image.setLeft(575);
    image.setTop(30);
    return;
  } else {
    console.log("Image URL is null");
  }
}

//Fetching Image URL
function fetchImageUrl(image_text) {
  // var prompt = "Suggest only one keyword for image to be searched on internet which needs to be inserted on a powerpoint slide that has following content - " + image_text;
  // keywords = getInfoHigh(prompt);
  // if (generatedUrls.includes(keywords)) {
  //     prompt = "Suggest only one keyword other than the following " + generatedUrls + " for image to be searched on internet which needs to be inserted on a powerpoint slide that has following content - " + image_text;
  //     keywords = getInfoHigh(prompt);
  //     generatedUrls.push(keywords);
  // }
  image_url = getImageUrlByKeywords(
    image_text,
    BING_API_KEY
  );
  console.log(image_url);
  return image_url;
}

//Get Information from API - Great Temperature
function getInfoHigh(prompt) {
  var response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/engines/text-davinci-002/completions",
    {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        Authorization:
          OPENAI_KEY,
      },
      payload: JSON.stringify({
        prompt: prompt,
        max_tokens: 1024,
        temperature: 1,
        n: 1,
      }),
    }
  );

  answer = JSON.parse(response).choices[0].text.trim();
  console.log(answer);
  return answer;
}

//Add Space after full stop
function addSpaceAfterFullStop(text) {
  var updatedText = text.replace(/\. /g, ".\n\n");

  return updatedText;
}

//removing extra stuff
function truncateString(input, count) {
  if (input.length <= count) {
    return input;
  }

  var truncatedText = input.substr(0, count);
  var lastFullStopIndex = truncatedText.lastIndexOf(".");

  if (lastFullStopIndex !== -1) {
    truncatedText = truncatedText.substr(0, lastFullStopIndex + 1);
  }
  console.log(truncatedText);
  return truncatedText;
}

//Creating a new slide
function createSlide(presentationId, num) {
  const pageId = Utilities.getUuid();

  const requests = [
    {
      createSlide: {
        objectId: pageId,
        insertionIndex: num,
        slideLayoutReference: {
          predefinedLayout: "BLANK",
        },
      },
    },
  ];

  try {
    const slide = Slides.Presentations.batchUpdate(
      {
        requests: requests,
      },
      presentationId
    );
    console.log(
      "Created Slide with ID: " + slide.replies[0].createSlide.objectId
    );
    return pageId;
  } catch (e) {
    // TODO (developer) - Handle Exception
    console.log("Failed with error %s", e.message);
  }
}

//Fetching image url from Bing API
function getImageUrlByKeywords(
  keywords,
  apiKey = BING_API_KEY
) {
  var query = encodeURIComponent(keywords);
  var url =
    "https://api.bing.microsoft.com/v7.0/images/search?q=" +
    query +
    "&count=1&safeSearch=Strict";

  var response = UrlFetchApp.fetch(url, {
    headers: {
      "Ocp-Apim-Subscription-Key": apiKey,
    },
  });

  var result = JSON.parse(response.getContentText());

  if (result.value && result.value.length > 0) {
    var bestMatchUrl = result.value[0].contentUrl;
    console.log(bestMatchUrl);
    return bestMatchUrl;
  } else {
    return null; // No image found
  }
}

//Generate keywords
function generateKeyWords(content) {
  jsonn = getInfo(
    "Understand and Generate 10 tools and computer technologies separated by comma related to the below project - " +
    content
  );
  console.log(jsonn);
  list = csvToList(jsonn);
  console.log(list);
}

function csvToList(csvString) {
  // Split the CSV string into an array of values
  if (csvString.includes(",")) {
    var csvValues = csvString.split(",");
  } else {
    var csvValues = csvString.split("\n");
  }

  // Trim whitespace from each value and create a new list
  var list = csvValues.map(function (value) {
    return removeNumbersSpacesFullStops(value.trim());
  });

  return list;
}

//String to list
function stringToJsonToList(inputString) {
  // Convert string to JSON object
  var jsonObject = JSON.parse(inputString);

  // Convert JSON object to list
  var list = [];
  for (var key in jsonObject) {
    if (jsonObject.hasOwnProperty(key)) {
      list.push(jsonObject[key]);
    }
  }

  return list;
}

function addRectangleToSlide(
  slide,
  left = 0,
  top = 0,
  width = 750,
  height = 10,
  color = "#FF0000"
) {
  var rectangle = slide.insertShape(SlidesApp.ShapeType.RECTANGLE);
  rectangle.setWidth(width);
  rectangle.setHeight(height);
  rectangle.setLeft(left);
  rectangle.setTop(top);
  rectangle.getFill().setSolidFill(color);
  rectangle.getBorder().getLineFill().setSolidFill(color);
  return;
}

function addTextToSlide(
  slide,
  left,
  top,
  width,
  height,
  content,
  fontSize,
  fontFamily,
  textColor
) {
  var textBox = slide.insertTextBox(content);
  textBox.setWidth(width);
  textBox.setHeight(height);
  textBox.setLeft(left);
  textBox.setTop(top);

  // Set text style for font size and family
  var textRange = textBox.getText();
  var textStyle = textRange.getTextStyle();
  textStyle.setFontSize(fontSize);
  textStyle.setFontFamily(fontFamily);
  textStyle.setBold(true);

  // Set text color
  textStyle.setForegroundColor(textColor);
}

function removeNumbersSpacesFullStops(input) {
  // Remove numbers
  var withoutNumbers = input.replace(/[0-9]/g, "");

  // Remove spaces
  var withoutSpaces = withoutNumbers.replace(/\s/g, "");

  // Remove full stops
  var withoutFullStops = withoutSpaces.replace(/\./g, "");

  return withoutFullStops;
}
