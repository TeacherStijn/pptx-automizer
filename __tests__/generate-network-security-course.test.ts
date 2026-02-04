import Automizer from '../src/automizer';

test('generate a network security course from JSON data', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const courseData = {
    courseTitle: 'Network Security Fundamentals',
    modules: [
      {
        number: 1,
        title: 'Threat Landscape & CIA Triad',
        icon: '🛡️',
        topics: [
          {
            title: 'Attack surfaces',
            content: 'Identify external, internal, and cloud exposure points.',
          },
          {
            title: 'Confidentiality, integrity, availability',
            content: 'Map security controls to each pillar of the CIA triad.',
          },
        ],
      },
      {
        number: 2,
        title: 'Network Defense Fundamentals',
        icon: '🧱',
        topics: [
          {
            title: 'Segmentation & zoning',
            content: 'Use VLANs and firewalls to limit lateral movement.',
          },
          {
            title: 'Secure protocols',
            content: 'Apply TLS, SSH, and IPsec to protect data in transit.',
          },
        ],
      },
      {
        number: 3,
        title: 'Monitoring & Detection',
        icon: '🔎',
        topics: [
          {
            title: 'Logging strategy',
            content: 'Centralize logs and define retention and alerting rules.',
          },
          {
            title: 'IDS/IPS',
            content: 'Deploy sensors and tune signatures for key threats.',
          },
        ],
      },
      {
        number: 4,
        title: 'Incident Response',
        icon: '🚨',
        topics: [
          {
            title: 'Playbooks',
            content: 'Create step-by-step guides for common incidents.',
          },
          {
            title: 'Containment & recovery',
            content: 'Isolate affected hosts and restore services safely.',
          },
        ],
      },
    ],
  };

  const pres = automizer
    .loadRoot('RootTemplate.pptx')
    .load('EmptySlide.pptx', 'empty');

  pres.addSlide('empty', 1, (slide) => {
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addText(courseData.courseTitle, {
        x: 0.6,
        y: 1.3,
        w: 12,
        h: 1.2,
        fontSize: 40,
        bold: true,
        color: '1F2937',
      });
      pptxGenJSSlide.addText('Course overview generated from JSON data', {
        x: 0.6,
        y: 2.6,
        w: 12,
        h: 0.6,
        fontSize: 18,
        color: '4B5563',
      });
    }, 'course title');
  });

  courseData.modules.forEach((module) => {
    pres.addSlide('empty', 1, (slide) => {
      slide.generate((pptxGenJSSlide) => {
        pptxGenJSSlide.addText(
          `${module.icon} Module ${module.number}: ${module.title}`,
          {
            x: 0.6,
            y: 0.6,
            w: 12,
            h: 0.8,
            fontSize: 28,
            bold: true,
            color: '111827',
          },
        );

        const topicLines = module.topics.map(
          (topic, index) =>
            `${index + 1}. ${topic.title} — ${topic.content}`,
        );

        pptxGenJSSlide.addText(topicLines.join('\n'), {
          x: 0.9,
          y: 1.8,
          w: 12,
          h: 4.8,
          fontSize: 18,
          color: '374151',
          valign: 'top',
        });
      }, `module-${module.number}`);
    });
  });

  await pres.write('network-security-course.example.pptx');
});
