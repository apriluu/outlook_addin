/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  const item = Office.context.mailbox.item;
  let resultsContainer = document.getElementById("results");
  resultsContainer.innerHTML = `<b>Subject:</b> ${item.subject}<br><br>`;

  try {
    // Get email body in HTML
    const htmlResult = await new Promise(resolve => 
      item.body.getAsync(Office.CoercionType.Html, resolve)
    );

    if (htmlResult.status !== Office.AsyncResultStatus.Succeeded) {
      throw new Error("Error reading email content");
    }

    const originalHtml = htmlResult.value;
    const input = document.getElementById("keywords").value;
    const keywords = input.split(",")
      .map(p => p.trim())
      .filter(p => p.length > 0);

    if (keywords.length === 0) {
      resultsContainer.innerHTML += "<p>Please enter keywords to filter.</p>";
      return;
    }

    // Process HTML to show only relevant sections
    const filteredHtml = filterHtml(originalHtml, keywords);
    
    // Display results
    if (filteredHtml.trim() === "") {
      resultsContainer.innerHTML += `<p>No sections found containing ALL keywords.</p>`;
    } else {
      resultsContainer.innerHTML += `
        <h3>Filtered email (showing only sections with ALL keywords):</h3>
        <div style="border: 1px solid #ddd; padding: 10px; max-height: 500px; overflow-y: auto;">
          ${filteredHtml}
        </div>
      `;
    }

  } catch (error) {
    resultsContainer.innerHTML += `<p style="color: red;">Error: ${error.message}</p>`;
  }
}

function filterHtml(html, keywords) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  
  function processNode(node) {
    // For text nodes, check if they contain ALL keywords
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent.toLowerCase();
      const containsAll = keywords.every(keyword => 
        text.includes(keyword.toLowerCase())
      );
      
      if (!containsAll) {
        return false;
      }
    }
    else if (node.nodeType === Node.ELEMENT_NODE) {
      let hasRelevantContent = false;
      const children = Array.from(node.childNodes);
      
      for (let i = children.length - 1; i >= 0; i--) {
        const child = children[i];
        const isRelevant = processNode(child);
        
        if (!isRelevant) {
          node.removeChild(child);
        } else {
          hasRelevantContent = true;
        }
      }
      
      // Highlight keywords in texts
      if (hasRelevantContent && node.childNodes.length > 0) {
        highlightKeywordsInNode(node, keywords);
        return true;
      }
      
      return hasRelevantContent;
    }
    
    return true;
  }
  
  const body = doc.body;
  processNode(body);
  
  return body.innerHTML;
}

function highlightKeywordsInNode(node, keywords) {
  // Only process text nodes
  if (node.nodeType === Node.TEXT_NODE) {
    let text = node.textContent;
    
    keywords.forEach(keyword => {
      const regex = new RegExp(`(${escapeRegExp(keyword)})`, 'gi');
      text = text.replace(regex, '<mark style="background-color: yellow;">$1</mark>');
    });
    
    const newNode = document.createRange().createContextualFragment(text);
    node.parentNode.replaceChild(newNode, node);
  } else if (node.nodeType === Node.ELEMENT_NODE) {
    Array.from(node.childNodes).forEach(child => {
      highlightKeywordsInNode(child, keywords);
    });
  }
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}