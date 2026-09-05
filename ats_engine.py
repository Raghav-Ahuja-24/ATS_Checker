"""
ATS Engine — Hardcoded Resume-to-Job Scorer
============================================
Implements the scoring methodology from "Ats Scorer.pdf".
No LLM / API key required. All scoring is deterministic and
based on text parsing, keyword matching, and weighted component scores.

Weights (from the PDF):
  Required Skills:          30%
  Relevant Experience:      20%
  Responsibility Match:     15%
  Preferred Skills:         10%
  Education:                 5%
  Certifications:            5%
  Job Title / Seniority:     5%
  Domain / Industry:         5%
  Keyword Coverage:          5%
"""

import json
import re
from difflib import SequenceMatcher

# ── Skill Normalization Map ──────────────────────────────────────────────────
SKILL_ALIASES = {
    "js": "javascript", "ts": "typescript", "py": "python",
    "postgres": "postgresql", "react.js": "react", "reactjs": "react",
    "node.js": "node", "nodejs": "node", "vue.js": "vue", "vuejs": "vue",
    "angular.js": "angular", "angularjs": "angular",
    "next.js": "nextjs", "nuxt.js": "nuxtjs",
    "express.js": "express", "expressjs": "express",
    "mongo": "mongodb", "k8s": "kubernetes",
    "aws ec2": "amazon ec2", "gcp": "google cloud",
    "c#": "csharp", "c++": "cplusplus",
    "ml": "machine learning", "ai": "artificial intelligence",
    "dl": "deep learning", "nlp": "natural language processing",
    "cv": "computer vision", "ops": "operations",
    "devops": "devops", "ci/cd": "cicd",
    "tf": "terraform", "docker compose": "docker-compose",
    "rest api": "rest", "restful": "rest",
    "graphql": "graphql", "sql server": "mssql",
    "ms sql": "mssql", "mysql": "mysql",
    ".net": "dotnet", "asp.net": "dotnet",
}

# Common tech-skill tokens that are too short for generic word extraction
SHORT_TECH_TOKENS = {
    "aws", "gcp", "sql", "css", "html", "api", "git", "ci", "cd",
    "npm", "php", "ios", "rpa", "sap", "erp", "crm", "etl", "bi",
    "vue", "svn", "ssh", "dns", "tcp", "udp", "xml", "json", "yaml",
    "jsx", "tsx", "go", "r", "c", "ui", "ux", "qa",
}

# Domain keywords for domain matching
DOMAIN_KEYWORDS = {
    "fintech": ["fintech", "financial", "banking", "payments", "trading"],
    "healthcare": ["healthcare", "medical", "clinical", "health", "pharma", "biotech"],
    "cybersecurity": ["cybersecurity", "security", "infosec", "penetration", "soc", "siem", "threat"],
    "telecom": ["telecom", "telecommunications", "5g", "networking", "voip"],
    "ecommerce": ["ecommerce", "e-commerce", "retail", "marketplace", "shopping"],
    "manufacturing": ["manufacturing", "industrial", "supply chain", "logistics"],
    "education": ["education", "edtech", "e-learning", "lms", "learning"],
    "cloud": ["cloud", "aws", "azure", "gcp", "saas", "iaas", "paas"],
}

# Education level hierarchy
EDUCATION_LEVELS = {
    "phd": 5, "doctorate": 5, "ph.d": 5,
    "master": 4, "mba": 4, "m.tech": 4, "m.sc": 4, "ms": 4, "m.s": 4,
    "mca": 4, "m.e": 4, "m.a": 4,
    "bachelor": 3, "b.tech": 3, "b.sc": 3, "b.e": 3, "bca": 3,
    "b.com": 3, "b.a": 3, "bba": 3, "degree": 3,
    "diploma": 2, "associate": 2,
    "high school": 1, "12th": 1, "hsc": 1,
}

# Seniority keywords
SENIORITY_LEVELS = {
    "intern": 1, "trainee": 1, "fresher": 1,
    "junior": 2, "associate": 2, "entry": 2,
    "mid": 3, "intermediate": 3,
    "senior": 4, "sr": 4, "lead": 4,
    "staff": 5, "principal": 5, "architect": 5,
    "manager": 5, "director": 6, "vp": 7, "head": 6,
    "cto": 8, "ceo": 8, "chief": 8, "executive": 7,
}

# ── Helper Functions ─────────────────────────────────────────────────────────

def _normalize_text(text):
    """Lowercase, collapse whitespace, strip."""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', str(text).lower().strip())


def _normalize_skill(skill):
    """Normalize a skill token using the alias map."""
    s = skill.strip().lower()
    return SKILL_ALIASES.get(s, s)


def _extract_skills(text):
    """
    Extract skill-like tokens from text.
    Returns a set of normalized skill strings.
    """
    text = _normalize_text(text)
    if not text:
        return set()

    skills = set()

    # Extract multi-word phrases first (2-3 word combos found between commas / pipes / semicolons)
    phrases = re.split(r'[,;|•·\n\r]+', text)
    for phrase in phrases:
        phrase = phrase.strip()
        if 2 <= len(phrase.split()) <= 4 and len(phrase) > 3:
            skills.add(_normalize_skill(phrase))

    # Extract individual words (length >= 3, or known short tokens)
    words = re.findall(r'[a-z][a-z0-9.#+\-]*[a-z0-9+#]|[a-z]{2,}', text)
    for w in words:
        nw = _normalize_skill(w)
        if len(nw) >= 3 or nw in SHORT_TECH_TOKENS:
            skills.add(nw)

    # Remove noise: section-header words and common non-skill words
    noise_words = {
        "required", "preferred", "skills", "experience", "education",
        "certification", "certifications", "responsibilities", "duties",
        "qualification", "qualifications", "mandatory", "essential",
        "minimum", "desired", "bonus", "nice", "have", "must",
        "good", "plus", "with", "years", "year", "looking", "seeking",
        "candidate", "ideal", "about", "role", "position", "company",
        "team", "work", "working", "ability", "strong", "excellent",
        "understanding", "knowledge", "familiar", "familiarity",
        "using", "used", "include", "including", "such", "like",
        "also", "well", "will", "able", "should", "shall",
        "domain", "industry", "requirements", "required skills:",
    }
    skills -= noise_words

    # Remove phrases that look like section headers
    skills = {s for s in skills if not any(
        s.startswith(prefix) for prefix in [
            "required skills:", "preferred skills:", "must have",
            "nice to have", "good to have",
        ]
    )}

    return skills


def _extract_years(text):
    """Extract numeric years-of-experience from text. Returns float or None."""
    text = _normalize_text(text)
    # Patterns like "5 years", "5+ years", "5-7 years", "5.5 years"
    patterns = [
        r'(\d+\.?\d*)\s*\+?\s*(?:years?|yrs?|yr)',
        r'(\d+\.?\d*)\s*-\s*\d+\.?\d*\s*(?:years?|yrs?|yr)',
        r'experience\s*[:\-]?\s*(\d+\.?\d*)',
    ]
    values = []
    for pat in patterns:
        matches = re.findall(pat, text)
        values.extend(float(m) for m in matches)
    return max(values) if values else None


def _detect_education_level(text):
    """Detect highest education level mentioned. Returns (level_int, label)."""
    text = _normalize_text(text)
    best_level = 0
    best_label = "unknown"
    for keyword, level in EDUCATION_LEVELS.items():
        if keyword in text and level > best_level:
            best_level = level
            best_label = keyword
    return best_level, best_label


def _detect_seniority(text):
    """Detect seniority level from text. Returns (level_int, label)."""
    text = _normalize_text(text)
    best_level = 0
    best_label = "unknown"
    for keyword, level in SENIORITY_LEVELS.items():
        if re.search(r'\b' + re.escape(keyword) + r'\b', text):
            if level > best_level:
                best_level = level
                best_label = keyword
    return best_level, best_label


def _detect_domains(text):
    """Detect domains mentioned in text. Returns list of domain names."""
    text = _normalize_text(text)
    detected = []
    for domain, keywords in DOMAIN_KEYWORDS.items():
        if any(kw in text for kw in keywords):
            detected.append(domain)
    return detected


def _fuzzy_match(a, b, threshold=0.75):
    """Return True if two strings are similar above threshold."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio() >= threshold


def _match_skills(required, candidate):
    """
    Match required skills against candidate skills.
    Returns (exact, equivalent, partial, missing) as lists.
    """
    exact = []
    equivalent = []
    partial = []
    missing = []

    for req in required:
        req_n = _normalize_skill(req)
        if req_n in candidate:
            exact.append(req)
        elif any(_fuzzy_match(req_n, cs, 0.85) for cs in candidate):
            equivalent.append(req)
        elif any(_fuzzy_match(req_n, cs, 0.6) for cs in candidate):
            partial.append(req)
        else:
            missing.append(req)

    return exact, equivalent, partial, missing


def _extract_requirements_from_jd(jd_text):
    """
    Parse a job description into structured requirement categories.
    Returns a dict with keys: required_skills, preferred_skills,
    experience_years, education, certifications, responsibilities, domains.
    """
    text = _normalize_text(jd_text)

    result = {
        "required_skills": [],
        "preferred_skills": [],
        "experience_years": None,
        "education": "",
        "certifications": [],
        "responsibilities": [],
        "domains": [],
    }

    # Split JD into sections by headings or bullet structure
    lines = jd_text.strip().split('\n')
    current_section = "general"

    required_markers = ["required", "must have", "mandatory", "essential", "minimum"]
    preferred_markers = ["preferred", "nice to have", "good to have", "bonus", "desired", "plus"]
    responsibility_markers = ["responsibilit", "duties", "you will", "role involves", "what you'll do"]
    education_markers = ["education", "qualification", "degree"]
    certification_markers = ["certification", "certified", "license"]

    for line in lines:
        line_lower = line.strip().lower()

        # Detect section headers
        if any(m in line_lower for m in required_markers):
            current_section = "required"
        elif any(m in line_lower for m in preferred_markers):
            current_section = "preferred"
        elif any(m in line_lower for m in responsibility_markers):
            current_section = "responsibilities"
        elif any(m in line_lower for m in education_markers):
            current_section = "education"
        elif any(m in line_lower for m in certification_markers):
            current_section = "certifications"

        # Extract skills from bullet points or comma-separated items
        stripped = re.sub(r'^[\-\*•·\d.)\]]+\s*', '', line.strip())
        if not stripped or len(stripped) < 3:
            continue

        if current_section == "required":
            # Extract individual skills from the line
            skills = _extract_skills(stripped)
            result["required_skills"].extend(skills)
        elif current_section == "preferred":
            skills = _extract_skills(stripped)
            result["preferred_skills"].extend(skills)
        elif current_section == "responsibilities":
            if len(stripped) > 10:
                result["responsibilities"].append(stripped)
        elif current_section == "education":
            result["education"] += " " + stripped
        elif current_section == "certifications":
            result["certifications"].append(stripped.lower())

    # Fallback: if no structured sections found, extract all skills as required
    if not result["required_skills"]:
        result["required_skills"] = list(_extract_skills(jd_text))

    # Extract experience requirement
    result["experience_years"] = _extract_years(jd_text)

    # Extract education
    if not result["education"]:
        result["education"] = jd_text  # let education detector scan full text

    # Detect domains
    result["domains"] = _detect_domains(jd_text)

    # Deduplicate
    result["required_skills"] = list(set(result["required_skills"]))
    result["preferred_skills"] = list(set(result["preferred_skills"]))
    result["certifications"] = list(set(result["certifications"]))

    return result


def _extract_candidate_profile(candidate_text):
    """
    Parse candidate resume/profile text into structured data.
    Returns a dict with keys: skills, experience_years, education_level,
    education_label, certifications, titles, domains, raw_text.
    """
    text = _normalize_text(candidate_text)

    profile = {
        "skills": _extract_skills(candidate_text),
        "experience_years": _extract_years(candidate_text),
        "education_level": 0,
        "education_label": "unknown",
        "certifications": [],
        "titles": [],
        "domains": _detect_domains(candidate_text),
        "seniority_level": 0,
        "seniority_label": "unknown",
        "raw_text": text,
    }

    # Education
    profile["education_level"], profile["education_label"] = _detect_education_level(candidate_text)

    # Seniority
    profile["seniority_level"], profile["seniority_label"] = _detect_seniority(candidate_text)

    # Certifications — look for cert-like lines
    cert_patterns = [
        r'((?:aws|azure|gcp|google|cisco|oracle|pmp|scrum|itil|'
        r'comptia|salesforce|vmware|red hat|rhce|ccna|ccnp|'
        r'certified|certification)[^\n,;]*)',
    ]
    for pat in cert_patterns:
        matches = re.findall(pat, text)
        profile["certifications"].extend(m.strip() for m in matches)
    profile["certifications"] = list(set(profile["certifications"]))

    # Job titles — look for designation / title fields
    title_patterns = [
        r'(?:designation|title|position|role)\s*[:\-]\s*([^\n,;]+)',
    ]
    for pat in title_patterns:
        matches = re.findall(pat, text)
        profile["titles"].extend(m.strip() for m in matches)

    return profile


# ── Component Scorers ────────────────────────────────────────────────────────

def _score_required_skills(jd_req, candidate):
    """Score: Required Skills (30% weight). Returns 0-100."""
    required = jd_req["required_skills"]
    if not required:
        return 100, [], [], [], []

    exact, equiv, partial, missing = _match_skills(required, candidate["skills"])

    total = len(required)
    score = 0
    if total > 0:
        # exact = 100%, equivalent = 85%, partial = 50%
        score = ((len(exact) * 1.0 + len(equiv) * 0.85 + len(partial) * 0.5) / total) * 100
        score = min(100, int(score))

    return score, exact, equiv, partial, missing


def _score_preferred_skills(jd_req, candidate):
    """Score: Preferred Skills (10% weight). Returns 0-100."""
    preferred = jd_req["preferred_skills"]
    if not preferred:
        return 100, [], []

    exact, equiv, partial, missing = _match_skills(preferred, candidate["skills"])
    total = len(preferred)
    matched_all = exact + equiv + partial

    score = 0
    if total > 0:
        score = ((len(exact) * 1.0 + len(equiv) * 0.85 + len(partial) * 0.5) / total) * 100
        score = min(100, int(score))

    return score, matched_all, missing


def _score_experience(jd_req, candidate):
    """Score: Relevant Experience (20% weight). Returns 0-100 and analysis dict."""
    required_yrs = jd_req["experience_years"]
    candidate_yrs = candidate["experience_years"]

    analysis = {
        "required_years": required_yrs,
        "candidate_years": candidate_yrs,
        "match": "UNKNOWN",
    }

    if required_yrs is None:
        analysis["match"] = "NOT_SPECIFIED"
        return 80, analysis  # No requirement stated — give decent score

    if candidate_yrs is None:
        analysis["match"] = "UNKNOWN"
        return 40, analysis  # Can't determine

    ratio = candidate_yrs / required_yrs if required_yrs > 0 else 1.0

    if ratio >= 1.0:
        analysis["match"] = "MEETS_OR_EXCEEDS"
        score = 100
    elif ratio >= 0.75:
        analysis["match"] = "PARTIALLY_MEETS"
        score = 75
    elif ratio >= 0.5:
        analysis["match"] = "BELOW_REQUIREMENT"
        score = 50
    else:
        analysis["match"] = "SIGNIFICANTLY_BELOW"
        score = 25

    return score, analysis


def _score_education(jd_req, candidate):
    """Score: Education (5% weight). Returns 0-100 and analysis dict."""
    jd_edu_level, jd_edu_label = _detect_education_level(jd_req["education"])
    cand_level = candidate["education_level"]
    cand_label = candidate["education_label"]

    analysis = {
        "required": jd_edu_label,
        "candidate": cand_label,
        "match": "UNKNOWN",
    }

    if jd_edu_level == 0:
        analysis["match"] = "NOT_SPECIFIED"
        return 80, analysis

    if cand_level == 0:
        analysis["match"] = "UNKNOWN"
        return 40, analysis

    if cand_level > jd_edu_level:
        analysis["match"] = "EXCEEDS"
        return 100, analysis
    elif cand_level == jd_edu_level:
        analysis["match"] = "MEETS"
        return 100, analysis
    elif cand_level == jd_edu_level - 1:
        analysis["match"] = "PARTIALLY_MEETS"
        return 60, analysis
    else:
        analysis["match"] = "DOES_NOT_MEET"
        return 20, analysis


def _score_certifications(jd_req, candidate):
    """Score: Certifications (5% weight). Returns 0-100 and analysis dict."""
    required_certs = jd_req["certifications"]
    candidate_certs = candidate["certifications"]

    analysis = {
        "required": required_certs,
        "candidate_has": candidate_certs,
        "match": "UNKNOWN",
    }

    if not required_certs:
        analysis["match"] = "NOT_SPECIFIED"
        return 80, analysis

    if not candidate_certs:
        analysis["match"] = "NO_CERTS_FOUND"
        return 20, analysis

    cand_text = " ".join(candidate_certs)
    matched = []
    missing_c = []
    for rc in required_certs:
        if any(_fuzzy_match(rc, cc, 0.6) for cc in candidate_certs):
            matched.append(rc)
        else:
            missing_c.append(rc)

    analysis["matched"] = matched
    analysis["missing"] = missing_c

    if not missing_c:
        analysis["match"] = "ALL_MET"
        return 100, analysis
    else:
        ratio = len(matched) / len(required_certs) if required_certs else 0
        analysis["match"] = "PARTIALLY_MET"
        return int(ratio * 100), analysis


def _score_job_title(jd_text, candidate):
    """Score: Job Title / Seniority (5% weight). Returns 0-100."""
    jd_seniority, jd_label = _detect_seniority(jd_text)
    cand_seniority = candidate["seniority_level"]
    cand_label = candidate["seniority_label"]

    if jd_seniority == 0:
        return 80  # No clear seniority in JD

    diff = abs(jd_seniority - cand_seniority)
    if diff == 0:
        return 100
    elif diff == 1:
        return 80
    elif diff == 2:
        return 55
    else:
        return 30


def _score_domain(jd_req, candidate):
    """Score: Domain / Industry (5% weight). Returns 0-100."""
    jd_domains = jd_req["domains"]
    cand_domains = candidate["domains"]

    if not jd_domains:
        return 80  # No specific domain required

    if not cand_domains:
        return 40  # Can't determine candidate's domain

    overlap = set(jd_domains) & set(cand_domains)
    if overlap:
        return 100
    else:
        return 30


def _score_responsibilities(jd_req, candidate):
    """Score: Responsibility Match (15% weight). Returns 0-100."""
    responsibilities = jd_req["responsibilities"]
    if not responsibilities:
        # No explicit responsibilities — score based on overall keyword overlap
        return 70

    resume_text = candidate["raw_text"]
    matches = 0
    for resp in responsibilities:
        resp_keywords = set(re.findall(r'\b[a-z]{4,}\b', resp.lower()))
        if not resp_keywords:
            continue
        resume_words = set(re.findall(r'\b[a-z]{4,}\b', resume_text))
        overlap = resp_keywords & resume_words
        if len(overlap) / len(resp_keywords) >= 0.4:
            matches += 1

    if not responsibilities:
        return 70

    ratio = matches / len(responsibilities)
    return min(100, int(ratio * 100))


def _score_keyword_coverage(jd_text, candidate_text):
    """Score: Keyword / Terminology Coverage (5% weight). Returns 0-100."""
    jd_words = set(re.findall(r'\b[a-z]{4,}\b', _normalize_text(jd_text)))
    resume_words = set(re.findall(r'\b[a-z]{4,}\b', _normalize_text(candidate_text)))

    # Remove very common stop words
    stop_words = {
        "with", "that", "this", "have", "from", "will", "been", "were",
        "they", "their", "about", "which", "would", "could", "should",
        "also", "into", "more", "other", "some", "such", "than", "then",
        "what", "when", "where", "your", "must", "able", "work",
        "well", "good", "make", "like", "just", "over", "only",
        "very", "each", "much", "both", "does", "most",
    }
    jd_words -= stop_words
    resume_words -= stop_words

    if not jd_words:
        return 80

    overlap = jd_words & resume_words
    ratio = len(overlap) / len(jd_words)
    return min(100, int(ratio * 100))


# ── Grammar Checker ──────────────────────────────────────────────────────────

# Common misspellings found in resumes (misspelled → correct)
_COMMON_MISSPELLINGS = {
    "acheive": "achieve", "acheivement": "achievement",
    "accross": "across", "adress": "address",
    "agile": None,  # correct — exclude
    "analsis": "analysis", "anaysis": "analysis",
    "bussiness": "business", "buisness": "business",
    "calender": "calendar", "carrer": "career",
    "collabrate": "collaborate", "colloborate": "collaborate",
    "commited": "committed", "communiation": "communication",
    "complience": "compliance", "concatination": "concatenation",
    "custmer": "customer", "databse": "database",
    "decison": "decision", "definately": "definitely",
    "dependancy": "dependency", "develope": "develop",
    "developement": "development", "diffrent": "different",
    "efficent": "efficient", "enviroment": "environment",
    "excellant": "excellent", "excercise": "exercise",
    "expirence": "experience", "experiance": "experience",
    "functionlity": "functionality", "guidence": "guidance",
    "hiearchy": "hierarchy", "immediatly": "immediately",
    "implmentation": "implementation", "independant": "independent",
    "infomation": "information", "infrastrucure": "infrastructure",
    "initally": "initially", "integeration": "integration",
    "knowlege": "knowledge", "langauge": "language",
    "liason": "liaison", "maintainance": "maintenance",
    "managment": "management", "manageing": "managing",
    "mileston": "milestone", "neccessary": "necessary",
    "occured": "occurred", "occurence": "occurrence",
    "optimzation": "optimization", "organistion": "organisation",
    "performace": "performance", "prioratize": "prioritize",
    "priviledge": "privilege", "proccess": "process",
    "profesional": "professional", "proficency": "proficiency",
    "progamming": "programming", "programing": "programming",
    "recomendation": "recommendation", "refrence": "reference",
    "relevent": "relevant", "reponsible": "responsible",
    "responsibile": "responsible", "scalibility": "scalability",
    "scheduele": "schedule", "seperate": "separate",
    "sevral": "several", "sofware": "software",
    "specifc": "specific", "stratagic": "strategic",
    "succesful": "successful", "sucess": "success",
    "sustainble": "sustainable", "sytem": "system",
    "technolgy": "technology", "thier": "their",
    "teh": "the", "transfered": "transferred",
    "utilisation": None,  # British spelling — not an error
    "wirless": "wireless", "writen": "written",
}


def _check_grammar(text):
    """
    Rule-based grammar checker for resume text.
    Returns a dict with:
        issues: list of {type, description, example} dicts
        issue_count: total number of issues found
        severity: 'none' | 'low' | 'medium' | 'high'
    """
    if not text or len(text.strip()) < 20:
        return {"issues": [], "issue_count": 0, "severity": "none"}

    issues = []
    original_text = str(text)

    # ── 1. Repeated consecutive words ("the the", "is is") ────────────────
    repeated = re.findall(r'\b(\w{2,})\s+\1\b', original_text, re.IGNORECASE)
    for word in repeated:
        issues.append({
            "type": "repeated_word",
            "description": f"Repeated word: '{word} {word}'",
            "example": f"{word} {word}",
        })

    # ── 2. Common misspellings ────────────────────────────────────────────
    words_in_text = re.findall(r'\b[a-z]{3,}\b', original_text.lower())
    seen_misspellings = set()
    for word in words_in_text:
        if word in _COMMON_MISSPELLINGS and _COMMON_MISSPELLINGS[word] is not None:
            if word not in seen_misspellings:
                seen_misspellings.add(word)
                issues.append({
                    "type": "misspelling",
                    "description": f"Possible misspelling: '{word}' → '{_COMMON_MISSPELLINGS[word]}'",
                    "example": word,
                })

    # ── 3. Uncapitalized sentence starts ─────────────────────────────────
    sentences = re.split(r'[.!?]\s+', original_text)
    uncap_count = 0
    for sent in sentences:
        sent = sent.strip()
        if sent and len(sent) > 3 and sent[0].islower() and not sent[0].isdigit():
            # Skip lines that look like key:value or bullet items
            if not re.match(r'^[\-•*]', sent) and ':' not in sent[:15]:
                uncap_count += 1
    if uncap_count > 0:
        issues.append({
            "type": "capitalization",
            "description": f"{uncap_count} sentence(s) start without capitalization",
            "example": f"{uncap_count} uncapitalized starts",
        })

    # ── 4. Subject-verb disagreement patterns ────────────────────────────
    sv_patterns = [
        (r'\b(I|we|they)\s+(is|was|has been)\b', "subject-verb disagreement"),
        (r'\b(he|she|it)\s+(are|were|have been)\b', "subject-verb disagreement"),
        (r'\b(I)\s+is\b', "subject-verb disagreement"),
        (r'\b(they|we)\s+is\b', "subject-verb disagreement"),
    ]
    for pattern, desc in sv_patterns:
        matches = re.findall(pattern, original_text, re.IGNORECASE)
        for match in matches:
            issues.append({
                "type": "grammar",
                "description": f"Possible {desc}: '{' '.join(match) if isinstance(match, tuple) else match}'",
                "example": ' '.join(match) if isinstance(match, tuple) else match,
            })

    # ── 5. Missing articles before singular nouns in common patterns ─────
    missing_article_patterns = [
        r'\b(am|is|was)\s+(senior|junior|lead|experienced|skilled)\s+[a-z]+er?\b',
    ]
    for pattern in missing_article_patterns:
        if re.search(pattern, original_text, re.IGNORECASE):
            match_text = re.search(pattern, original_text, re.IGNORECASE).group()
            # Check if preceded by an article
            pos = original_text.lower().find(match_text.lower())
            prefix = original_text[max(0, pos-5):pos].strip().lower()
            if prefix not in ('a', 'an', 'the'):
                issues.append({
                    "type": "missing_article",
                    "description": f"Possible missing article before: '{match_text}'",
                    "example": match_text,
                })

    # ── 6. Inconsistent tense (mixing past and present in bullet points) ─
    past_verbs = len(re.findall(
        r'\b(managed|developed|implemented|designed|created|built|led|'
        r'delivered|maintained|improved|achieved|worked|handled|coordinated|'
        r'established|reduced|increased|launched|migrated|deployed|authored|'
        r'architected|mentored|resolved|automated|streamlined)\b',
        original_text, re.IGNORECASE
    ))
    present_verbs = len(re.findall(
        r'\b(manage|develop|implement|design|create|build|lead|'
        r'deliver|maintain|improve|achieve|handle|coordinate|'
        r'establish|reduce|increase|launch|migrate|deploy|author|'
        r'architect|mentor|resolve|automate|streamline)\b',
        original_text, re.IGNORECASE
    ))
    if past_verbs > 2 and present_verbs > 2:
        ratio = min(past_verbs, present_verbs) / max(past_verbs, present_verbs)
        if ratio > 0.4:  # Significant mix
            issues.append({
                "type": "tense_inconsistency",
                "description": f"Inconsistent verb tense: {past_verbs} past-tense vs {present_verbs} present-tense action verbs",
                "example": f"{past_verbs} past / {present_verbs} present",
            })

    # ── 7. Excessive use of first person ──────────────────────────────────
    first_person_count = len(re.findall(r'\bI\b', original_text))
    if first_person_count > 5:
        issues.append({
            "type": "style",
            "description": f"Excessive first-person pronoun usage ({first_person_count} occurrences of 'I') — resumes typically avoid this",
            "example": f"'I' used {first_person_count} times",
        })

    # ── Determine severity ────────────────────────────────────────────────
    count = len(issues)
    if count == 0:
        severity = "none"
    elif count <= 2:
        severity = "low"
    elif count <= 5:
        severity = "medium"
    else:
        severity = "high"

    return {
        "issues": issues,
        "issue_count": count,
        "severity": severity,
    }


# ── Mandatory Requirement Penalty ────────────────────────────────────────────

def _apply_mandatory_penalty(score, missing_required, total_required):
    """
    Apply mandatory requirement penalty per the PDF:
    If mandatory requirements are not met, apply substantial penalty.
    """
    if total_required == 0:
        return score, []

    miss_ratio = len(missing_required) / total_required
    penalties = []

    if miss_ratio >= 0.5:
        penalty = 25
        penalties.append(f"Major penalty (-{penalty}): {len(missing_required)}/{total_required} required skills missing")
        score = max(0, score - penalty)
    elif miss_ratio >= 0.3:
        penalty = 15
        penalties.append(f"Moderate penalty (-{penalty}): {len(missing_required)}/{total_required} required skills missing")
        score = max(0, score - penalty)
    elif miss_ratio > 0:
        penalty = 8
        penalties.append(f"Minor penalty (-{penalty}): {len(missing_required)}/{total_required} required skills missing")
        score = max(0, score - penalty)

    return score, penalties


def _apply_grammar_penalty(score, grammar_result):
    """
    Apply grammar penalty based on issues found.
    Max penalty: -10 points.
    """
    penalties = []
    severity = grammar_result["severity"]
    count = grammar_result["issue_count"]

    if severity == "none":
        return score, penalties

    if severity == "high":
        penalty = 10
    elif severity == "medium":
        penalty = 5
    else:  # low
        penalty = 2

    penalties.append(
        f"Grammar penalty (-{penalty}): {count} issue(s) found — {severity} severity"
    )
    score = max(0, score - penalty)

    return score, penalties


# ── Score Band ───────────────────────────────────────────────────────────────

def _score_band(score):
    if score >= 90:
        return "Excellent"
    elif score >= 80:
        return "Strong"
    elif score >= 70:
        return "Good"
    elif score >= 60:
        return "Moderate"
    elif score >= 40:
        return "Weak"
    else:
        return "Poor"


# ── Public API ───────────────────────────────────────────────────────────────

def init_gemini(api_key):
    """(Deprecated) No-op — kept for backward compatibility."""
    pass


def run_deep_ats_scan(api_key, job_description, candidate_text):
    """
    Hardcoded ATS scorer following the Ats Scorer.pdf methodology.
    Parses the job description and candidate text, then scores across
    9 weighted components. No LLM or API key required.

    Args:
        api_key: Ignored (kept for backward compatibility with app.py).
        job_description: The job description text.
        candidate_text: The candidate resume/profile text (or JSON string).

    Returns:
        dict matching the expected ATS JSON output format.
    """
    # If candidate_text is JSON (from app.py), flatten it to a string
    try:
        cdata = json.loads(candidate_text)
        if isinstance(cdata, dict):
            candidate_text = " ".join(
                f"{k}: {v}" for k, v in cdata.items()
                if v and not str(k).startswith("_")
            )
    except (json.JSONDecodeError, TypeError):
        pass

    # ── Step 1: Extract structured data ──────────────────────────────────
    jd_req = _extract_requirements_from_jd(job_description)
    candidate = _extract_candidate_profile(candidate_text)

    # ── Step 2: Score each component ─────────────────────────────────────
    req_score, exact_skills, equiv_skills, partial_skills, missing_skills = \
        _score_required_skills(jd_req, candidate)

    pref_score, pref_matched, pref_missing = \
        _score_preferred_skills(jd_req, candidate)

    exp_score, exp_analysis = _score_experience(jd_req, candidate)

    resp_score = _score_responsibilities(jd_req, candidate)

    edu_score, edu_analysis = _score_education(jd_req, candidate)

    cert_score, cert_analysis = _score_certifications(jd_req, candidate)

    title_score = _score_job_title(job_description, candidate)

    domain_score = _score_domain(jd_req, candidate)

    keyword_score = _score_keyword_coverage(job_description, candidate_text)

    # ── Step 3: Weighted combination ─────────────────────────────────────
    weighted_score = (
        req_score   * 0.30 +
        exp_score   * 0.20 +
        resp_score  * 0.15 +
        pref_score  * 0.10 +
        edu_score   * 0.05 +
        cert_score  * 0.05 +
        title_score * 0.05 +
        domain_score * 0.05 +
        keyword_score * 0.05
    )
    ats_score = int(weighted_score)

    # ── Step 4: Mandatory requirement penalty ────────────────────────────
    total_required = len(jd_req["required_skills"])
    ats_score, penalties = _apply_mandatory_penalty(ats_score, missing_skills, total_required)

    # ── Step 4b: Grammar check & penalty ─────────────────────────────────
    grammar_result = _check_grammar(candidate_text)
    ats_score, grammar_penalties = _apply_grammar_penalty(ats_score, grammar_result)
    penalties.extend(grammar_penalties)

    ats_score = max(0, min(100, ats_score))

    # ── Step 5: Build mandatory requirements classification ──────────────
    mandatory_met = exact_skills + equiv_skills
    mandatory_partial = partial_skills
    mandatory_not_met = missing_skills

    # ── Step 6: Identify strengths & weaknesses ──────────────────────────
    strengths = []
    weaknesses = []

    if req_score >= 70:
        strengths.append(f"Strong required-skill coverage ({req_score}%)")
    elif req_score < 40:
        weaknesses.append(f"Low required-skill coverage ({req_score}%)")

    if exp_score >= 80:
        strengths.append(f"Experience meets/exceeds requirement")
    elif exp_score < 50:
        weaknesses.append(f"Experience below requirement")

    if edu_score >= 80:
        strengths.append(f"Education meets requirement ({edu_analysis['match']})")
    elif edu_score < 50:
        weaknesses.append(f"Education gap ({edu_analysis['match']})")

    if keyword_score >= 60:
        strengths.append(f"Good keyword/terminology overlap ({keyword_score}%)")
    elif keyword_score < 40:
        weaknesses.append(f"Low keyword overlap ({keyword_score}%)")

    if domain_score >= 80:
        strengths.append("Domain experience aligns")
    elif domain_score < 50:
        weaknesses.append("Limited domain overlap")

    if pref_score >= 70:
        strengths.append(f"Good preferred-skill coverage ({pref_score}%)")

    if resp_score >= 70:
        strengths.append(f"Responsibility alignment ({resp_score}%)")
    elif resp_score < 40:
        weaknesses.append(f"Low responsibility match ({resp_score}%)")

    if grammar_result["severity"] == "none":
        strengths.append("Clean writing — no grammar issues detected")
    elif grammar_result["severity"] in ("medium", "high"):
        weaknesses.append(f"Grammar issues detected ({grammar_result['issue_count']} issues, {grammar_result['severity']} severity)")

    # ── Step 7: Build evidence list ──────────────────────────────────────
    evidence = []
    if exact_skills:
        evidence.append(f"Exact skill matches: {', '.join(exact_skills[:8])}")
    if equiv_skills:
        evidence.append(f"Equivalent skill matches: {', '.join(equiv_skills[:5])}")
    if partial_skills:
        evidence.append(f"Partial skill matches: {', '.join(partial_skills[:5])}")
    if missing_skills:
        evidence.append(f"Missing required skills: {', '.join(missing_skills[:8])}")
    evidence.append(f"Experience: {exp_analysis.get('candidate_years', 'unknown')} yrs "
                     f"(required: {exp_analysis.get('required_years', 'not specified')})")
    evidence.append(f"Education: candidate={edu_analysis['candidate']}, "
                     f"required={edu_analysis['required']} → {edu_analysis['match']}")

    # ── Step 8: Build summary ────────────────────────────────────────────
    summary = (
        f"ATS heuristic scan scored {ats_score}/100 ({_score_band(ats_score)}). "
        f"Matched {len(exact_skills)+len(equiv_skills)} of "
        f"{total_required} required skills "
        f"({len(partial_skills)} partial). "
        f"Experience: {exp_analysis['match']}. "
        f"Education: {edu_analysis['match']}."
    )

    # ── Step 9: Confidence ───────────────────────────────────────────────
    # Higher when we have more data points to work with
    data_points = sum([
        1 if jd_req["required_skills"] else 0,
        1 if jd_req["experience_years"] is not None else 0,
        1 if candidate["experience_years"] is not None else 0,
        1 if candidate["education_level"] > 0 else 0,
        1 if candidate["skills"] else 0,
        1 if jd_req["responsibilities"] else 0,
    ])
    confidence = min(95, 40 + data_points * 10)

    # ── Build output JSON ────────────────────────────────────────────────
    return {
        "ats_score": ats_score,
        "score_band": _score_band(ats_score),
        "summary": summary,
        "component_scores": {
            "required_skills": req_score,
            "relevant_experience": exp_score,
            "responsibility_match": resp_score,
            "preferred_skills": pref_score,
            "education": edu_score,
            "certifications": cert_score,
            "job_title_seniority": title_score,
            "domain_match": domain_score,
            "keyword_coverage": keyword_score,
        },
        "mandatory_requirements": {
            "met": mandatory_met[:10],
            "partially_met": mandatory_partial[:10],
            "not_met": mandatory_not_met[:10],
            "unknown": [],
        },
        "matched_skills": (exact_skills + equiv_skills)[:15],
        "missing_skills": missing_skills[:15],
        "partial_matches": partial_skills[:10],
        "experience_analysis": exp_analysis,
        "education_analysis": edu_analysis,
        "certification_analysis": cert_analysis,
        "responsibility_analysis": {
            "score": resp_score,
            "jd_responsibilities_count": len(jd_req["responsibilities"]),
        },
        "grammar_analysis": grammar_result,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "penalties": penalties,
        "evidence": evidence,
        "confidence": confidence,
    }
