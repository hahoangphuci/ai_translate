// home.js - Home page JavaScript
document.addEventListener("DOMContentLoaded", function () {
  initializeHomePage();
});

function initializeHomePage() {
  setupSmoothScrolling();
  setupPricingButtons();
  setupHeroAnimations();
  setupMobileMenu();
}

// Global helper used by pricing CTA buttons
function goToUpgrade(plan) {
  const selected = String(plan || "")
    .trim()
    .toLowerCase();
  const token = localStorage.getItem("token") || "";

  // If already logged in, go straight to dashboard and auto-open upgrade flow
  if (token) {
    if (selected && selected !== "free") {
      window.location.href = `/dashboard?upgrade_plan=${encodeURIComponent(selected)}&autocreate=1`;
    } else {
      window.location.href = "/dashboard";
    }
    return;
  }

  // Not logged in: remember intended plan for after login
  if (selected && selected !== "free") {
    try {
      localStorage.setItem("pending_upgrade_plan", selected);
    } catch (e) {
      // ignore
    }
  }

  window.location.href = "/auth";
}

function setupMobileMenu() {
  // Mobile menu toggle
  window.toggleMenu = function () {
    const navLinks = document.querySelector(".nav-links");
    const hamburger = document.querySelector(".hamburger");

    navLinks.classList.toggle("mobile-menu");
    hamburger.classList.toggle("active");
  };

  // Close mobile menu when clicking outside or on a link
  document.addEventListener("click", function (e) {
    const navLinks = document.querySelector(".nav-links");
    const hamburger = document.querySelector(".hamburger");

    if (
      !e.target.closest(".nav-container") &&
      navLinks.classList.contains("mobile-menu")
    ) {
      navLinks.classList.remove("mobile-menu");
      hamburger.classList.remove("active");
    }
  });

  // Close mobile menu when clicking on nav links
  document.querySelectorAll(".nav-links a").forEach((link) => {
    link.addEventListener("click", function () {
      const navLinks = document.querySelector(".nav-links");
      const hamburger = document.querySelector(".hamburger");

      navLinks.classList.remove("mobile-menu");
      hamburger.classList.remove("active");
    });
  });
}

function setupSmoothScrolling() {
  // Smooth scrolling for navigation links
  document.querySelectorAll('a[href^="#"]').forEach((anchor) => {
    anchor.addEventListener("click", function (e) {
      e.preventDefault();
      const target = document.querySelector(this.getAttribute("href"));
      if (target) {
        target.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }
    });
  });
}

function scrollToFeatures() {
  const featuresSection = document.getElementById("features");
  if (featuresSection) {
    featuresSection.scrollIntoView({
      behavior: "smooth",
      block: "start",
    });
  }
}

function setupPricingButtons() {
  // Pricing CTA buttons are wired via onclick="goToUpgrade(...)" in home.html.
  // Keep this function for backward compatibility, but don't override CTA behavior.
}

function setupHeroAnimations() {
  // Add intersection observer for animations
  const observerOptions = {
    threshold: 0.1,
    rootMargin: "0px 0px -50px 0px",
  };

  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add("animate-in");
        // Animate stats numbers
        if (entry.target.classList.contains("stat-item")) {
          animateNumber(entry.target.querySelector(".stat-number"));
        }
      }
    });
  }, observerOptions);

  // Observe elements for animation
  document
    .querySelectorAll(".feature-card, .pricing-card, .benefit-card, .stat-item")
    .forEach((card) => {
      observer.observe(card);
    });
}

function animateNumber(element) {
  const target = parseInt(element.getAttribute("data-target"));
  const duration = 2000; // 2 seconds
  const step = target / (duration / 16); // 60fps
  let current = 0;

  const timer = setInterval(() => {
    current += step;
    if (current >= target) {
      element.textContent = target.toLocaleString();
      clearInterval(timer);
    } else {
      element.textContent = Math.floor(current).toLocaleString();
    }
  }, 16);
}

// Add some CSS animations dynamically
const style = document.createElement("style");
style.textContent = `
    .feature-card, .pricing-card, .benefit-card {
        opacity: 0;
        transform: translateY(30px);
        transition: opacity 0.6s ease, transform 0.6s ease;
    }

    .feature-card.animate-in, .pricing-card.animate-in, .benefit-card.animate-in {
        opacity: 1;
        transform: translateY(0);
    }

    .hero-content {
        animation: fadeInUp 1s ease-out;
    }

    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
`;
document.head.appendChild(style);

function quickTranslate() {
  const text = document.getElementById("quick-text").value.trim();
  const targetLang = document.getElementById("quick-target-lang").value;
  const resultDiv = document.getElementById("quick-result");
  const outputSpan = document.getElementById("quick-output");

  if (!text) {
    showQuickMessage("Vui lòng nhập văn bản cần dịch!", "error");
    return;
  }

  // Simple demo translations for common phrases
  const translations = {
    hello: { vi: "xin chào", en: "hello", fr: "bonjour", de: "hallo" },
    "how are you": {
      vi: "bạn thế nào",
      en: "how are you",
      fr: "comment allez-vous",
      de: "wie geht es dir",
    },
    "thank you": { vi: "cảm ơn", en: "thank you", fr: "merci", de: "danke" },
    "good morning": {
      vi: "chào buổi sáng",
      en: "good morning",
      fr: "bonjour",
      de: "guten morgen",
    },
    goodbye: {
      vi: "tạm biệt",
      en: "goodbye",
      fr: "au revoir",
      de: "auf wiedersehen",
    },
  };

  // Try to find exact match first
  const lowerText = text.toLowerCase();
  let translation = translations[lowerText]?.[targetLang];

  // If no exact match, try partial match
  if (!translation) {
    for (const [key, value] of Object.entries(translations)) {
      if (lowerText.includes(key)) {
        translation = value[targetLang];
        break;
      }
    }
  }

  // If still no translation, provide a generic response
  if (!translation) {
    if (targetLang === "vi") {
      translation = "[Dịch sang tiếng Việt: " + text + "]";
    } else if (targetLang === "en") {
      translation = "[Translated to English: " + text + "]";
    } else if (targetLang === "fr") {
      translation = "[Traduit en français: " + text + "]";
    } else if (targetLang === "de") {
      translation = "[Übersetzt auf Deutsch: " + text + "]";
    }
  }

  outputSpan.textContent = translation;
  resultDiv.style.display = "block";

  // Hide result after 10 seconds
  setTimeout(() => {
    resultDiv.style.display = "none";
  }, 10000);
}

function showQuickMessage(message, type = "info") {
  // Simple notification for quick translate
  const notification = document.createElement("div");
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: ${type === "error" ? "rgba(220, 53, 69, 0.9)" : "rgba(40, 167, 69, 0.9)"};
    color: white;
    padding: 15px 20px;
    border-radius: 8px;
    z-index: 10000;
    font-weight: 500;
  `;
  notification.textContent = message;
  document.body.appendChild(notification);

  setTimeout(() => {
    document.body.removeChild(notification);
  }, 3000);
}
