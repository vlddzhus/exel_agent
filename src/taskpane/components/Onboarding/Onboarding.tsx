import React, { useState } from "react";
import { makeStyles, tokens, Button, Text } from "@fluentui/react-components";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    padding: "24px",
    textAlign: "center",
    gap: "16px",
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  desc: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    lineHeight: 1.5,
  },
  items: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
  },
  dots: {
    display: "flex",
    gap: "6px",
  },
  dot: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    backgroundColor: tokens.colorNeutralStroke1,
    border: "none",
    cursor: "pointer",
    padding: 0,
  },
  dotActive: {
    backgroundColor: tokens.colorBrandBackground,
  },
});

interface OnboardingProps {
  onDone: () => void;
}

const slides = [
  { title: ru.onboarding.slide1.title, desc: ru.onboarding.slide1.desc },
  { title: ru.onboarding.slide2.title, items: ru.onboarding.slide2.items },
  {
    title: ru.onboarding.slide3.title,
    desc: ru.onboarding.slide3.desc,
    cta: ru.onboarding.slide3.cta,
  },
];

export const Onboarding: React.FC<OnboardingProps> = ({ onDone }) => {
  const styles = useStyles();
  const [slide, setSlide] = useState(0);
  const s = slides[slide];

  return (
    <div className={styles.root}>
      <div className={styles.title}>{s.title}</div>
      {"desc" in s && <div className={styles.desc}>{s.desc}</div>}
      {"items" in s && s.items && (
        <div className={styles.items}>
          {s.items.map((item, i) => (
            <div key={i}>{item}</div>
          ))}
        </div>
      )}
      <div className={styles.dots}>
        {slides.map((_, i) => (
          <button
            key={i}
            className={`${styles.dot} ${i === slide ? styles.dotActive : ""}`}
            onClick={() => setSlide(i)}
            aria-label={`Слайд ${i + 1}`}
          />
        ))}
      </div>
      {slide < slides.length - 1 ? (
        <Button appearance="primary" onClick={() => setSlide(slide + 1)}>
          Далее
        </Button>
      ) : (
        <Button appearance="primary" onClick={onDone}>
          {"cta" in s ? s.cta : "Начать"}
        </Button>
      )}
    </div>
  );
};
