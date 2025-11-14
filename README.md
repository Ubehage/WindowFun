# WindowFun

**WindowFun** is a classic VB6 "trolling" and stress-test application, designed purely for fun and experimentation with window management techniques.  
It was created late at night as a playful project. Expect chaos if you start closing windows!

## What Does WindowFun Do?

- Pops up a window at a random position and color.
- Each window’s **title bar** shows its “level” (used internally by the app).
- The **main text inside the window** shows how many windows are currently open within that level.
- When you close a window, _all windows in that level_ are closed. For each closed window, two new windows appear at a higher level.
- **Ending the program** (by clicking or using a close effect) will close **all open windows, in all levels.**

Great for stress-testing your system’s window handling, or just for a harmless prank!

## Safety First

Upon starting, WindowFun presents a warning and explanatory dialog.  
You must confirm before the main application starts.

**Warning:**  
WindowFun can spawn large numbers of windows very rapidly.  
Do NOT run in environments where this could cause problems or with unsaved work open.

## Usage & Controls

Every window displays available controls in its main text area:

- **Click (inside window):** Closes _all_ windows and exits the program.
- **ESC:** Fade the window to black, then close.
- **F:** Shrink the window, then close.
- **C:** Fade out and close.
- **A:** Perform all closing effects.
- **B:** Toggle beep when closing: ON/OFF.

The **window's title bar** shows the current “level” (internal app logic).  
The **main text** reports how many windows are open at that level.

## Technical Notes

- Fully self-contained: all required sources included, no external DLL dependencies.
- Written in VB6, with a focus on API calls and custom controls (no runtime modules).
- Code is optimized for high speed and responsive behavior—no lag tolerated!

## Code Quality & Philosophy

- Code is intentionally messy and experimental, as a result of rapid prototyping for fun.
- Created for enjoyment, stress testing, and late-night learning.
- Criticism and suggestions are welcome; feel free to fork, refactor, or improve!
- No guarantees or warranties.
##
- The codebase contains **no comments**. This is an intentional choice and reflects my personal coding style.
- WindowFun was written for fun, and has been optimized as flaws were spotted during use.
- What you see is what you get: a late-night project, quick and messy, with improvements added along the way.
- You are welcome to interpret, refactor, or document the code further—but I prefer it as-is.
## License

MIT License.

Use at your own risk, and only in safe environments.

## Contributing

Pull requests for bug fixes, optimization, or code cleanups are encouraged.
You're welcome to add features, document quirks, or customize your own version.

---

**Enjoy WindowFun responsibly!**
