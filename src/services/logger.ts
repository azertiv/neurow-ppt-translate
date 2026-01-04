export type LogLevel = "info" | "warn" | "error" | "dim";

export class Logger {
  private el: HTMLElement;
  constructor(el: HTMLElement) {
    this.el = el;
  }

  clear() {
    this.el.innerHTML = "";
  }

  log(message: string, level: LogLevel = "info") {
    const line = document.createElement("div");
    line.className = `logLine ${level}`;
    const t = new Date().toLocaleTimeString();
    line.textContent = `${t} — ${message}`;
    this.el.appendChild(line);
    this.el.scrollTop = this.el.scrollHeight;
  }
}
