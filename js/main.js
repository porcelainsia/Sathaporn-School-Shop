import { mount } from './runtime.js';
import { Component } from './app-logic.js';
mount(Component, document.getElementById('app-template'), document.getElementById('root'));
