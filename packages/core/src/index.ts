export interface PresentationElement {
  type: string;
}

export interface PresentationSlideInput {
  id?: string;
  elements?: PresentationElement[];
}

export interface PresentationInit {
  title?: string;
}

export interface PresentationSlide {
  id: string;
  elements: PresentationElement[];
}

export interface PresentationDocument {
  title?: string;
  slides: PresentationSlide[];
  addSlide(input?: PresentationSlideInput): PresentationSlide;
}

function createSlideId(index: number): string {
  return `slide-${index + 1}`;
}

export function createPresentation(init: PresentationInit = {}): PresentationDocument {
  const slides: PresentationSlide[] = [];
  const document: PresentationDocument = {
    slides,
    addSlide(input: PresentationSlideInput = {}): PresentationSlide {
      const slide: PresentationSlide = {
        id: input.id ?? createSlideId(slides.length),
        elements: [...(input.elements ?? [])],
      };

      slides.push(slide);
      return slide;
    },
  };

  if (init.title !== undefined) {
    document.title = init.title;
  }

  return document;
}
