// ==UserScript==
// @name         Microsoft Teams Emoji Extension
// @namespace    http://tampermonkey.net/
// @description  Add easy to access WebEx style emojis to Microsoft Teams
// @author       Luca Magni
// @version      1.0.0
// @updateURL    https://raw.githubusercontent.com/m4gni/teams/main/teams.user.js
// @downloadURL  https://raw.githubusercontent.com/m4gni/teams/main/teams.user.js
// @match        https://teams.live.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    var style = document.createElement('style');
    style.innerHTML = `.icon-container {display: flex; justify-content: space-between; margin-bottom: 10px;} .icon {width: 100%; height: 45px; display: flex; justify-content: center; align-items: center; background: #292929; margin-right: 10px; border: 1px solid #525252; border-radius: 6px; font-size: 120%; transition: 0.3s ease;} .icon:hover {margin-top: -2.5px; cursor: pointer; box-shadow: 0 5px 10px -4px rgba(0, 0, 0, 0.3);} .icon:last-child {margin-right: 0;}`;
    document.head.appendChild(style);

    function waitForElementToExist(selector, callback) {
        const intervalId = setInterval(() => {
            const parentDiv = document.getElementById('ExperienceContainerManagerMountElement');
            if (parentDiv) {
                const iframe = parentDiv.querySelector('iframe');
                if (iframe) {
                    const iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
                    const chatButton = iframeDocument.getElementById('chat-button');
                    if (chatButton) {
                        clearInterval(intervalId);
                        callback(iframe, chatButton);
                    }
                }
            }
        }, 100);
    }

    waitForElementToExist('#ExperienceContainerManagerMountElement iframe', (iframe, chatButton) => {
        chatButton.click();
        setTimeout(() => {
            function sendEmoji(type) {
                const parentDiv = document.getElementById('ExperienceContainerManagerMountElement');
                const iframe = parentDiv.querySelector('iframe');
                const iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
                const contentEditableDiv = iframeDocument.querySelector('div[data-tid="ckeditor"]');
                const paragraphElement = contentEditableDiv.querySelector('p');
                paragraphElement.innerHTML = '';

                const emojiHTMLMap = {
                    thumbs_up: '<readonly contenteditable="false" title="Yes" itemid="yes" itemtype="http://schema.skype.com/Emoji" itemscope="👍" aria-label="Yes"><img src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/yes/default/50_f.png?v=v70" alt="Yes" draggable="false" width="20px" height="20px" aria-label="Yes" style="overflow: hidden; margin: 0px 1px;"></readonly>',
                    thumbs_down: '<readonly contenteditable="false" title="No" itemid="no" itemtype="http://schema.skype.com/Emoji" itemscope="👎" aria-label="No"><img src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/no/default/50_f.png?v=v50" alt="No" draggable="false" width="20px" height="20px" aria-label="No" style="overflow: hidden; margin: 0px 1px;"></readonly>',
                    tick_yes: '<readonly contenteditable="false" title="Tick button" itemid="2705_whiteheavycheckmark" itemtype="http://schema.skype.com/Emoji" itemscope="✅" aria-label="Tick button"><img src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/2705_whiteheavycheckmark/default/50_f.png?v=v14" alt="Tick button" draggable="false" width="20px" height="20px" aria-label="Tick button" style="overflow: hidden; margin: 0px 1px;"></readonly>',
                    tick_no: '<readonly contenteditable="false" title="Cross mark button" itemid="274e_negativesquaredcrossmark" itemtype="http://schema.skype.com/Emoji" itemscope="❎" aria-label="Cross mark button"><img src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/274e_negativesquaredcrossmark/default/50_f.png?v=v11" alt="Cross mark button" draggable="false" width="20px" height="20px" aria-label="Cross mark button" style="overflow: hidden; margin: 0px 1px;"></readonly>'
                };

                if (emojiHTMLMap.hasOwnProperty(type)) {
                    const tempDiv = iframeDocument.createElement('div');
                    tempDiv.innerHTML = emojiHTMLMap[type];
                    const newContent = tempDiv.firstChild;
                    paragraphElement.appendChild(newContent);
                    const sendButton = iframeDocument.querySelector('[data-tid="newMessageCommands-send"]');
                    sendButton.click();
                }
            }

            const parentDiv = document.getElementById('ExperienceContainerManagerMountElement');
            const iframe = parentDiv.querySelector('iframe');
            const iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
            const uiBoxWithAttribute = iframeDocument.querySelector('.ui-box div[data-tid="compose-bottom-toolbar"]').closest('.ui-box');

            const htmlToInsert = '<div class="icon-container">'+
                '<div class="icon" id="thumbs_up">👍</div>'+
                '<div class="icon" id="thumbs_down">👎</div>'+
                '<div class="icon" id="tick_yes">✅</div>'+
                '<div class="icon" id="tick_no">❎</div>'+
                '</div>';

            const newDiv = iframeDocument.createElement('div');
            newDiv.innerHTML = htmlToInsert;

            if (uiBoxWithAttribute) {
                uiBoxWithAttribute.insertBefore(newDiv, uiBoxWithAttribute.firstChild);

                iframeDocument.getElementById('thumbs_up').addEventListener('click', () => {
                    sendEmoji('thumbs_up');
                });
                iframeDocument.getElementById('thumbs_down').addEventListener('click', () => {
                    sendEmoji('thumbs_down');
                });
                iframeDocument.getElementById('tick_yes').addEventListener('click', () => {
                    sendEmoji('tick_yes');
                });
                iframeDocument.getElementById('tick_no').addEventListener('click', () => {
                    sendEmoji('tick_no');
                });
            }
        }, 1000);
    });
})();
