/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */

let speechFlag = false;

Office.onReady(async function (info) {
  displayStartingMessage("Hi! I'm NoviWord, your Word assistant bot. I can help you create documents, modify content, and insert useful information seamlessly. How can I assist you today?");
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
document.getElementById("sendButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    displayChatMessage(question, '', "User",directLine1);
      await getBotResponse(directLine1, question);
  }
};

document.getElementById("userInput").addEventListener("keydown", async function (event) {
  if (event.key === "Enter") {
    // Check if Enter key is pressed
    event.preventDefault(); 
    const question = document.getElementById("userInput").value;
    if (question) {
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
  mic.classList.toggle("recording");
  const popup = window.open('speech.html', 'SpeechRecognition', 'width=1,height=1');
  speechFlag = true;
  // Listen for messages from the pop-up window
  window.addEventListener("message", async function eventHandler(event){
      if (event.origin !== window.location.origin) return; // Security check
      // Get the recognized text from the pop-up
      transcript = event.data;
      // Insert recognized text into user input
      console.log(transcript);
      document.getElementById("userInput").value = transcript;
      var question = document.getElementById("userInput").value  ;
    if (question) {
        document.getElementById("userInput").value ="";
        displayChatMessage(question, '', "User");
        await getBotResponse(directLine1, question);
      }
      popup.close();
      mic.classList.toggle("recording");
      window.removeEventListener("message", eventHandler);
  }, { once: true });
});
}
});

function displayStartingMessage(starter) {
  const chatWindow = document.getElementById("chatWindow");
  
  chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${starter}</div>`; 
    
}
// Display user question and bot response in chat window
function displayChatMessage(question, response, role,directLine) {
  const chatWindow = document.getElementById("chatWindow");
  // Check if response is valid and if attachments exist
  // eslint-disable-next-line no-constant-condition
  if (response && response.attachments && response.attachments.length > 0 && false) {
    response.attachments.forEach((attachment) => {
      
      if (attachment.content && attachment.content.buttons && attachment.content.buttons.length > 0) {
        attachment.content.buttons.forEach((button) => {
          if (button.type === "signin") {
            const signinButton = document.createElement("button");
            signinButton.innerText = button.title || "Sign In"; 
            signinButton.classList.add("ms-Button", "ms-Button--primary");
            signinButton.onclick = () => {
              window.open(button.value, "_blank"); 
            };
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${attachment.content.text}</div>`;
            chatWindow.appendChild(signinButton); 
          }
        });
      }
    });
  } else {
    // Regular message display if no attachments
    if (role === "bot") {
      if(response.speak==="Generate"){
        insertResponseIntoDocument(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>`; 
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("S.O.W. content generated in document");
        });
        speechFlag = false;  
        }
      }else if(response.speak==="Table"){

        insertResponseIntoDocumentAtCursor(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">Table has been generated in document</div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("Table has been generated in document");
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
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">Requested changes have been made in the document</div>`;  
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("Requested changes have been made in the document");
        });
        speechFlag = false;  
        } 
      }
      else if(response.speak==="interim"){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`;
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(response.text);
        });
        }      
      }
    
      else if(response.text){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`;
        document.getElementById("insertButton").style.display = "block";
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
        document.getElementById("insertButton").style.display = "none";
        chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;      
      }
    }
  }
  scrollToBottom();
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
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
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
  }, 100);
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
  let voices = window.speechSynthesis.getVoices();
  let femaleVoice = voices.find(voice => voice.name.includes("Female") ||
  voice.name.includes("Google UK English Female") ||
   voice.name.includes("Microsoft Zira")||
   voice.name.includes("Samantha")
  );
  const speech = new SpeechSynthesisUtterance(text);
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