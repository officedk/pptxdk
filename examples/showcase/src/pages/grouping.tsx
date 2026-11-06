import { CardsPage } from "../templates/cards-page.tsx";

export function GroupingPage() {
  return (
    <CardsPage
      slideId="grouping"
      title="Order from fragments."
      subtitle="Lorem ipsum dolor sit amet, consectetur adipiscing elit."
      items={[
        {
          title: "Lorem ipsum",
          body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
        },
        {
          title: "Dolor sit",
          body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
        },
        {
          title: "Amet consectetur",
          body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
        },
      ]}
    />
  );
}

export function RevealPage() {
  return (
    <CardsPage
      slideId="reveal"
      title="One idea at a time."
      subtitle="Lorem ipsum dolor sit amet, consectetur adipiscing elit."
      items={[
        {
          title: "Lorem ipsum",
          body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
        },
        {
          title: "Dolor sit amet",
          body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
        },
      ]}
    />
  );
}
