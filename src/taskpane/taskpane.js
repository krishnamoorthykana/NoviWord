/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */

const botUrl= "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview";
const welcomeMsg = "Hi, I am your word assistant bot-NoviWord";
const changeMsg = "Requested changes have been made in the document";
const sowMsg = "SOW content generated in document";
const tableMsg = "Table has been generated in document";
const botMessageTemplate = (message) => `
  <div class="bot-wrapper">
    <img width="20" height="20" src="../../assets/copilot.png"/> NoviWord
  </div>
  <div class="message bot">${message}</div>
`;
const userMessageTemplate = (message) => `
  <div class="user-wrapper">You</div>
  <div class="message user">${message}</div>
`;
let speechFlag = false;

Office.onReady(async function (info) {
  displayStartingMessage(welcomeMsg);
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
  //let flag=true;
 
document.getElementById("askButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    document.getElementById("headerId").style.display = "none";
    displayChatMessage(question, '', "User",directLine1);
      await getBotResponse(directLine1, question);
   
 
  }
};
 
document.getElementById("userInput").addEventListener("keydown", async function (event) {
  if (event.key === "Enter") {
    // Check if Enter key is pressed
    event.preventDefault(); // Prevents the default behavior (like submitting a form)
 
    const question = document.getElementById("userInput").value;
    if (question) {
      //document.getElementById("headerId").style.display = "none";
        displayChatMessage(question, '', "User",directLine1);
      await getBotResponse(directLine1, question);
     
  }
}});
 
// Handle the Insert button click
document.getElementById("insertButton").onclick = async function () {
  const response = document.getElementById("chatWindow").lastChild
    ? document.getElementById("chatWindow").lastChild.innerText
    : "";
  if (response) {
    await insertResponseIntoDocument(response);
  }
};
 
document.getElementById('startSpeechButton').addEventListener('click', function () {
  // Open a pop-up window to handle the speech
  const popup = window.open('speech.html', 'SpeechRecognition', 'width=400,height=300');
  speechFlag = true;
  // Listen for messages from the pop-up window
  window.addEventListener("message", async function eventHandler(event){
      if (event.origin !== window.location.origin) return; // Security check
 
      // Get the recognized text from the pop-up
      transcript = event.data;
 
      // Insert recognized text into user input
      console.log(transcript);
      document.getElementById("userInput").value = transcript;
      var question = document.getElementById("userInput").value ;
    if (question) {
        document.getElementById("userInput").innerText="";
        displayChatMessage(question, '', "User");
        await getBotResponse(directLine1, question);
      }
      popup.close();
      window.removeEventListener("message", eventHandler);
  }, { once: true });
});
 
}
});
 
function displayStartingMessage(starter) {
  const chatWindow = document.getElementById("chatWindow");
 
  chatWindow.innerHTML += botMessageTemplate(starter);
   
}
 
 
// Display user question and bot response in chat window
function displayChatMessage(question, response, role,directLine) {
  const chatWindow = document.getElementById("chatWindow");
 
  // Check if response is valid and if attachments exist
  // eslint-disable-next-line no-constant-condition
  if (response && response.attachments && response.attachments.length > 0 && false) {
    response.attachments.forEach((attachment) => {
      // Check if attachment content has 'buttons' and 'signin' type
      if (attachment.content && attachment.content.buttons && attachment.content.buttons.length > 0) {
        attachment.content.buttons.forEach((button) => {
          if (button.type === "signin") {
            // Create a sign-in button
            const signinButton = document.createElement("button");
            signinButton.innerText = button.title || "Sign In"; // Default title to "Sign In"
            signinButton.classList.add("ms-Button", "ms-Button--primary");
 
            // Open the sign-in URL when the button is clicked
            signinButton.onclick = () => {
              window.open(button.value, "_blank"); // Open the sign-in URL in a new tab
            };
 
            // Display the bot's message
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${attachment.content.text}</div>`;
            chatWindow.appendChild(signinButton); // Add the button after the message
          }
        });
      }
    });
  } else {
    // Regular message display if no attachments
    if (role === "bot") {
      if(response.speak==="Generate"){
 
        insertResponseIntoDocument(response.text);
        chatWindow.innerHTML += botMessageTemplate(sowMsg);
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(sowMsg);
        });
       
        speechFlag = false;  
        }
      }else if(response.speak==="Table"){
 
        insertResponseIntoDocumentAtCursor(response.text);
        chatWindow.innerHTML += botMessageTemplate(tableMsg);     
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(tableMsg);
        });
       
        speechFlag = false;  
        }
      }
      else if(response.speak==="Replace"){
        splitText=response.text
        textArray=splitText.split("|");
        replaceText(textArray[0],textArray[1]);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">Replaced ${textArray[0]} with ${textArray[1]} </div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(`Replaced ${textArray[0]} with ${textArray[1]}`);
        });
       
        speechFlag = false;  
        }
      }
      else if(response.speak==="Selected"){
        console.log("fetching selected data")
        getSelectedText(directLine);  
      }
      else if(response.speak==="paragraph"){
        setSelectedText(response.text);
        chatWindow.innerHTML += botMessageTemplate(changeMsg);  
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(changeMsg);
        });
       
        speechFlag = false;  
        }
      }
      else if(response.speak==="interim"){
        chatWindow.innerHTML += botMessageTemplate(response.text);
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(response.text);
        });
       
        }      
      }
   
      else if(response.text){
        chatWindow.innerHTML += botMessageTemplate(response.text);
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(response.text);
        });
       
        speechFlag = false;  
        }      
      }
     
    }
    else {
      if(question){
        chatWindow.innerHTML += userMessageTemplate(question);      
      }
     
    }
  }
  scrollToBottom();
  // Clear the input field
  document.getElementById("userInput").value = "";
}
 
// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.end);
    await context.sync();
  });
}
 
async function insertResponseIntoDocumentAtCursor(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.after);
    await context.sync();
  });
}
const initializeDirectLine = async function () {
  try {
    const response = await fetch(
      botUrl
    );
    const data = await response.json();
   
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
   
 
    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }
 
    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: "Hi",
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );
 
    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      console.log("Role", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Bot Response: ", activity.text);
        displayChatMessage(false, activity, activity.from.role,directLine);
       
      }
    });
    return directLine;
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
 
const getBotResponse = async function (directLine, question) {
  directLine
    .postActivity({
      from: { id: "10", name: "User" },
      type: "message",
      text: question,
    })
    .subscribe(             // calls the subscription already created in InitializeDirectline method
      (id) => console.log("Message sent with ID:", id),
      (error) => console.error("Error sending message:", error)
    );
 
}
function scrollToBottom() {
  const chatWindow = document.getElementById("chatWindow");
  setTimeout(() => {
    chatWindow.scrollTop = chatWindow.scrollHeight;
  }, 100); // Timeout ensures scroll happens after the new message is rendered
}
 
async function replaceText(oldText,NewText){
  await Word.run(async (context) => {
    let results = context.document.body.search(oldText);
    results.load();
    await context.sync();
   
    results.items.forEach(item => {
        item.insertText(NewText, Word.InsertLocation.replace);
    });
   
    await context.sync();
});
}
 
async function getSelectedText(directLine) {
  await Word.run(async (context) => {
    let range = context.document.getSelection();
    range.load("text");
    await context.sync();
    SelText=range.text;
    await getBotResponse(directLine, SelText);
});
 
}
 
async function setSelectedText(response) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(response, Word.InsertLocation.replace);
    await context.sync();
  });
}
 
 
 
function speakText(text) {
  console.log("Testing Text to Speech");
  let voices = window.speechSynthesis.getVoices();
  console.log("Voices******", voices);
  let femaleVoice = voices.find(voice => voice.name.includes("Female") ||
  voice.name.includes("Google UK English Female") ||
   voice.name.includes("Microsoft Zira")||
   voice.name.includes("Samantha")
  );
  console.log("Set voice********", femaleVoice);
  const speech = new SpeechSynthesisUtterance(text);
  // speech.lang = 'en-US'; // Set language
  // speech.rate = 1; // Speed of speech (0.1 to 10)
  // speech.pitch = 1; // Pitch (0 to 2)
  // speech.volume = 1; // Volume (0 to 1)
  if (femaleVoice) {
    speech.voice = femaleVoice;
} else {
    console.warn("Female voice not found. Using default voice.");
}
  window.speechSynthesis.speak(speech);
}
 
 
// Load voices properly before calling the function
function ensureVoicesLoaded(callback) {
  let voices = window.speechSynthesis.getVoices();
 
  if (voices.length > 0) {
      callback();
  } else {
      window.speechSynthesis.onvoiceschanged = callback;
  }
}
 