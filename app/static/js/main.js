document.addEventListener("DOMContentLoaded", () => {
    // Auto hide alerts
    document.querySelectorAll(".alert").forEach(alert => {
      setTimeout(() => { alert.classList.remove("show"); }, 4000);
    });
  
    // Button hover animation
    document.querySelectorAll("button, .btn").forEach(btn => {
      btn.addEventListener("mouseenter", () => {
        btn.style.transform = "scale(1.05)";
        btn.style.transition = "0.2s";
      });
      btn.addEventListener("mouseleave", () => { btn.style.transform = "scale(1)"; });
    });
  });
 