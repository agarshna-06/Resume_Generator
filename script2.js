document.addEventListener("DOMContentLoaded", function () {
    const form = document.getElementById("resume-form");
    const outputContainer = document.getElementById("output-container");
    const addProjectBtn = document.getElementById("add-project");

    // Handle adding new project entries
    addProjectBtn.addEventListener("click", function() {
        const projectsSection = this.previousElementSibling;
        const newProjectEntry = projectsSection.cloneNode(true);
        // Clear the values
        newProjectEntry.querySelectorAll('input, textarea').forEach(input => input.value = '');
        this.parentElement.insertBefore(newProjectEntry, this);
    });

    // Form Submission Handler
    form.addEventListener("submit", function (event) {
        event.preventDefault();
        const formData = new FormData(form);

        const resumeData = {
            name: formData.get("name"),
            degree: formData.get("degree"),
            graduation_year: formData.get("graduation_year"),
            tagline: formData.get("tagline"),
            email: formData.get("email"),
            phone: formData.get("phone"),
            linkedin: formData.get("linkedin"),
            github: formData.get("github"),
            education: {
                college: {
                    name: formData.get("college_name"),
                    degree: formData.get("college_degree"),
                    duration: formData.get("college_duration")
                },
                school: {
                    name: formData.get("school_name"),
                    program: formData.get("school_program"),
                    duration: formData.get("school_duration"),
                    score: formData.get("school_score")
                }
            },
            courses: [
                formData.get("course1"),
                formData.get("course2"),
                formData.get("course3"),
                formData.get("course4"),
                formData.get("course5"),
                formData.get("course6")
            ].filter(Boolean),
            experience: {
                title: formData.get("job_title"),
                company: formData.get("company"),
                duration: formData.get("duration"),
                description: formData.get("job_description"),
                technologies: formData.get("technologies")
            },
            projects: Array.from(form.querySelectorAll('.project-entry')).map(entry => ({
                title: entry.querySelector('[name="project_title"]').value,
                description: entry.querySelector('[name="project_description"]').value,
                technologies: entry.querySelector('[name="project_technologies"]').value
            })).filter(project => project.title),
            achievements: formData.get("achievements").split('\n').filter(Boolean),
            skills: [
                formData.get("skill1"),
                formData.get("skill2"),
                formData.get("skill3"),
                formData.get("skill4"),
                formData.get("skill5"),
                formData.get("skill6")
            ].filter(Boolean)
        };

        generateDOCX(resumeData);
    });

    // Function to generate DOCX
    window.generateDOCX = async function(resumeData) {
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    // Header with Name and Degree
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: resumeData.name,
                                bold: true,
                                size: 32,
                                color: "#2c7da0"
                            }),
                            new docx.TextRun({
                                text: `\n${resumeData.degree}, ${resumeData.graduation_year}`,
                                size: 24,
                                color: "#666666"
                            }),
                            new docx.TextRun({
                                text: `\n${resumeData.tagline}`,
                                size: 20,
                                color: "#666666",
                                italics: true
                            })
                        ],
                        spacing: { after: 300 },
                    }),

                    // Contact Information
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: "📧 ", size: 20 }),
                            new docx.TextRun({ text: resumeData.email, size: 20 }),
                            new docx.TextRun({ text: "   📱 ", size: 20 }),
                            new docx.TextRun({ text: resumeData.phone, size: 20 }),
                            new docx.TextRun({ text: "\n" }),
                            new docx.TextRun({ text: "🔗 ", size: 20 }),
                            new docx.TextRun({ text: resumeData.linkedin, size: 20 }),
                            new docx.TextRun({ text: "   💻 ", size: 20 }),
                            new docx.TextRun({ text: resumeData.github, size: 20 })
                        ],
                        spacing: { after: 400 },
                    }),

                    // Education Section
                    new docx.Paragraph({
                        text: "EDUCATION",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: `${resumeData.education.college.name}\n`,
                                bold: true,
                                size: 24,
                            }),
                            new docx.TextRun({
                                text: `${resumeData.education.college.degree}\n`,
                                size: 22,
                            }),
                            new docx.TextRun({
                                text: resumeData.education.college.duration,
                                size: 20,
                                color: "#666666"
                            })
                        ],
                        spacing: { after: 200 },
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: `${resumeData.education.school.name}\n`,
                                bold: true,
                                size: 24,
                            }),
                            new docx.TextRun({
                                text: `${resumeData.education.school.program}\n`,
                                size: 22,
                            }),
                            new docx.TextRun({
                                text: `${resumeData.education.school.duration} | ${resumeData.education.school.score}%`,
                                size: 20,
                                color: "#666666"
                            })
                        ],
                        spacing: { after: 300 },
                    }),

                    // Relevant Courses Section
                    new docx.Paragraph({
                        text: "RELEVANT COURSES",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    new docx.Paragraph({
                        children: resumeData.courses.map(course => 
                            new docx.TextRun({
                                text: `${course}   `,
                                size: 22,
                            })
                        ),
                        spacing: { after: 300 },
                    }),

                    // Experience Section
                    new docx.Paragraph({
                        text: "EXPERIENCE",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: `${resumeData.experience.title} - ${resumeData.experience.company}\n`,
                                bold: true,
                                size: 24,
                            }),
                            new docx.TextRun({
                                text: resumeData.experience.duration + "\n",
                                size: 20,
                                color: "#666666"
                            }),
                            new docx.TextRun({
                                text: resumeData.experience.description + "\n",
                                size: 22,
                            }),
                            new docx.TextRun({
                                text: "Technologies used: " + resumeData.experience.technologies,
                                size: 20,
                                color: "#666666"
                            })
                        ],
                        spacing: { after: 300 },
                    }),

                    // Projects Section
                    new docx.Paragraph({
                        text: "SELECTED PROJECTS",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    ...resumeData.projects.map(project => 
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: `${project.title}\n`,
                                    bold: true,
                                    size: 24,
                                }),
                                new docx.TextRun({
                                    text: project.description + "\n",
                                    size: 22,
                                }),
                                new docx.TextRun({
                                    text: "Technologies used: " + project.technologies,
                                    size: 20,
                                    color: "#666666"
                                })
                            ],
                            spacing: { after: 200 },
                        })
                    ),

                    // Achievements Section
                    new docx.Paragraph({
                        text: "ACHIEVEMENTS",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    ...resumeData.achievements.map(achievement => 
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: "• " + achievement,
                                    size: 22,
                                })
                            ],
                            spacing: { after: 100 },
                        })
                    ),

                    // Skills Section
                    new docx.Paragraph({
                        text: "SKILLS",
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { before: 200, after: 200 },
                        thematicBreak: true,
                    }),
                    new docx.Paragraph({
                        children: resumeData.skills.map(skill => 
                            new docx.TextRun({
                                text: `${skill}   `,
                                size: 22,
                            })
                        ),
                        spacing: { after: 200 },
                    }),
                ],
            }],
        });

        const blob = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${resumeData.name.replace(' ', '_')}_Resume.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
});