import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('ipcRenderer', {
     send: (channel, ...args) => ipcRenderer.send(channel, ...args),
     on: (channel, listener) => ipcRenderer.on(channel, listener),
     once: (channel, listener) => ipcRenderer.once(channel, listener),
     removeListener: (channel, listener) => ipcRenderer.removeListener(channel, listener),
     invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args)
});