
import Handsontable from 'handsontable/base';
import { TextEditor } from 'handsontable/editors/textEditor';


export class CustomEditor extends TextEditor {
  
  constructor(props: Handsontable.Core) {
    super(props);
  }

  override createElements() {
    super.createElements();

    this.TEXTAREA = document.createElement('input');
    this.TEXTAREA.setAttribute('placeholder', 'Custom placeholder');
    this.TEXTAREA.setAttribute('data-hot-input', 'true');
    this.textareaStyle = this.TEXTAREA.style;
    this.TEXTAREA_PARENT.innerText = '';
    this.TEXTAREA_PARENT.appendChild(this.TEXTAREA);
  }
}