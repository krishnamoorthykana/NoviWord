/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
  let directLine1=null;
  let flag = true;
  document.getElementById("myButton").onclick = function() {
    // Open the task pane when the button is clicked
    Office.context.ui.displayDialogAsync('https://krishnamoorthykana.github.io/NoviWord/src/taskpane/index.html', { height: 30, width: 20 });
  };
  document.addEventListener("DOMContentLoaded", async function () {
    if (flag) {
      directLine1 = await initializeDirectLine();
      flag = false;
    }
  });
   
  document.getElementById("askButton").onclick = async function () {
    const question = document.getElementById("userInput").value;
    if (question) {
     
      displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
     
    }
  };
   
  document.getElementById("userInput").addEventListener("keydown", async function (event) {
    if (event.key === "Enter") {
      // Check if Enter key is pressed
      event.preventDefault(); // Prevents the default behavior (like submitting a form)
   
      const question = document.getElementById("userInput").value;
      if (question) {
       
        displayChatMessage(question, '', "User");
        await getBotResponse(directLine1, question);
      }
    }
  });
   
  // Handle the Insert button click
  document.getElementById("insertButton").onclick = async function () {
    const response = document.getElementById("chatWindow").lastChild
      ? document.getElementById("chatWindow").lastChild.innerText
      : "";
    if (response) {
      await insertResponseIntoDocument(response);
    }
  };
  }
  });
   
   
  // Display user question and bot response in chat window
  function displayChatMessage(question, response, role) {
    const chatWindow = document.getElementById("chatWindow");
   
    // Check if response is valid and if attachments exist
    if (response && response.attachments && response.attachments.length > 0) {
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
              chatWindow.innerHTML += `<div class="bot"><img src="assets/copilot.png" alt="Copilot Icon" /> <br>${attachment.content.text}</div>`;
              chatWindow.appendChild(signinButton); // Add the button after the message
            }
          });
        }
      });
    } else {
      // Regular message display if no attachments
      if (role === "bot") {
        if(response.text){
          chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`;      }
      } else {
        if(question){
       
          chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;      }
       
      }
    }
   
    // Clear the input field
    document.getElementById("userInput").value = "";
  }
   
  // Function to insert the response into the Word document
  async function insertResponseIntoDocument(response) {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText(response, Word.InsertLocation.end);
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
          displayChatMessage(false, activity, activity.from.role);
         
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
   
   